const STATE = {
  executor: "",
  requests: [],
  filteredRequests: [],
  activeFilter: "ALL",
  selectedRequest: null,
  categories: [],
  cities: [],
  statuses: [],
  employees: []
};


// ============================================================
// START
// ============================================================

Office.onReady(async info => {

  initEvents();

  if (info.host !== Office.HostType.Excel) {
    return;
  }

  try {

    await loadDictionaries();

  } catch (error) {

    console.error(error);

  }

});


// ============================================================
// EVENTS
// ============================================================

function initEvents() {

  document
    .getElementById("currentExecutor")
    ?.addEventListener(
      "change",
      async event => {

        STATE.executor =
          event.target.value;

        await loadRequests();

      }
    );


  document
    .getElementById("searchInput")
    ?.addEventListener(
      "input",
      applyFilters
    );


  document
    .querySelectorAll(".filter-button")
    .forEach(button => {

      button.addEventListener(
        "click",
        () => {

          document
            .querySelectorAll(".filter-button")
            .forEach(item =>
              item.classList.remove("active")
            );


          button.classList.add("active");


          STATE.activeFilter =
            button.dataset.filter;


          applyFilters();

        }
      );

    });


  document
    .getElementById("newRequestButton")
    ?.addEventListener(
      "click",
      openCreateModal
    );


  document
    .getElementById("navNewRequest")
    ?.addEventListener(
      "click",
      openCreateModal
    );


  document
    .getElementById("refreshButton")
    ?.addEventListener(
      "click",
      loadRequests
    );


  document
    .getElementById("navRefresh")
    ?.addEventListener(
      "click",
      loadRequests
    );


  document
    .getElementById("saveCreateButton")
    ?.addEventListener(
      "click",
      createRequest
    );


  document
    .getElementById("saveStatusButton")
    ?.addEventListener(
      "click",
      saveStatusChange
    );


  document
    .getElementById("saveEditButton")
    ?.addEventListener(
      "click",
      saveEdit
    );


  document
    .getElementById("statusNew")
    ?.addEventListener(
      "change",
      handlePauseReason
    );


  document
    .querySelectorAll("[data-close]")
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


    categoryRange.load("values");
    cityRange.load("values");
    employeeRange.load("values");
    slaRange.load("values");


    await context.sync();


    STATE.categories =
      categoryRange.values
        .map(row => String(row[0] || "").trim())
        .filter(value => value);


    STATE.cities =
      cityRange.values
        .map(row => String(row[0] || "").trim())
        .filter(value => value);


    STATE.employees =
      employeeRange.values
        .map(row => String(row[0] || "").trim())
        .filter(value => value);


    STATE.statuses =
      slaRange.values
        .map(row => String(row[0] || "").trim())
        .filter(value => value);

  });


  fillSelect(
    "currentExecutor",
    STATE.employees,
    "Оберіть виконавця"
  );


  fillSelect(
    "createExecutor",
    STATE.employees,
    "Оберіть виконавця"
  );


  fillSelect(
    "createCategory",
    STATE.categories,
    "Оберіть категорію"
  );


  fillSelect(
    "editCategory",
    STATE.categories,
    "Оберіть категорію"
  );


  fillSelect(
    "createCity",
    STATE.cities,
    "Оберіть місто"
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
// LOAD REQUESTS
// ============================================================

async function loadRequests() {

  STATE.requests = [];
  STATE.filteredRequests = [];

  if (!STATE.executor) {

    renderRequests();

    return;

  }


  showLoading();


  try {

    STATE.requests =
      await Excel.run(async context => {

        const tableName =
          getEmployeeTableName(
            STATE.executor
          );


        const table =
          context.workbook.tables.getItem(
            tableName
          );


        table.rows.load("items");

        const headerRange =
          table.getHeaderRowRange();

        headerRange.load("values");


        await context.sync();


        const headers =
          headerRange.values[0];


        if (table.rows.items.length === 0) {
          return [];
        }


        const dataRange =
          table.getDataBodyRange();

        dataRange.load("values");


        await context.sync();


        return dataRange.values.map(row =>
          rowToRequest(
            headers,
            row
          )
        );

      });


    applyFilters();

  } catch (error) {

    console.error(error);

    showTableError(
      getErrorText(error)
    );

  }

}


// ============================================================
// ROW → OBJECT
// ============================================================

function rowToRequest(
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

    sd:
      String(
        get("Номер заявки SD")
      ).trim(),

    category:
      String(
        get("Категорія")
      ).trim(),

    description:
      String(
        get("Опис")
      ).trim(),

    city:
      String(
        get("Місто")
      ).trim(),

    address:
      String(
        get("Адреса / локація")
      ).trim(),

    created:
      get("Дата створення"),

    amount:
      get("Сума витрат"),

    customer:
      String(
        get("Замовник")
      ).trim(),

    status:
      String(
        get("Статус")
      ).trim(),

    executor:
      String(
        get("Виконавець")
      ).trim(),

    id:
      String(
        get("ID")
      ).trim(),

    source:
      String(
        get("Джерело заявки")
      ).trim(),

    statusSince:
      get("Поточний статус з"),

    previousStatus:
      String(
        get("Попередній статус")
      ).trim(),

    sla:
      Number(
        get("SLA, год")
      ) || 0,

    timeInStatus:
      Number(
        get("Час у статусі, год")
      ) || 0,

    slaOverdue:
      String(
        get("Прострочено SLA")
      ).trim(),

    plannedDate:
      get("Планова дата завершення"),

    dateOverdue:
      String(
        get("Прострочено по даті")
      ).trim(),

    pauseReason:
      String(
        get("Причина призупинення")
      ).trim(),

    comment:
      String(
        get("Коментар")
      ).trim()

  };

}


// ============================================================
// FILTERS
// ============================================================

function applyFilters() {

  const search =
    String(
      document
        .getElementById("searchInput")
        ?.value || ""
    )
      .trim()
      .toLowerCase();


  STATE.filteredRequests =
    STATE.requests.filter(request => {

      if (
        !passesStatusFilter(
          request
        )
      ) {
        return false;
      }


      if (!search) {
        return true;
      }


      const haystack = [
        request.id,
        request.sd,
        request.description,
        request.city,
        request.customer,
        request.category,
        request.status
      ]
        .join(" ")
        .toLowerCase();


      return haystack.includes(
        search
      );

    });


  updateCounters();
  renderRequests();

}


// ============================================================

function passesStatusFilter(
  request
) {

  switch (
    STATE.activeFilter
  ) {

    case "NEW":

      return request.status === "Нова";


    case "PAUSED":

      return request.status === "Призупинена";


    case "OVERDUE":

      return (
        request.slaOverdue === "Так" ||
        request.dateOverdue === "Так"
      );


    case "WORK":

      return (
        request.status !== "Нова" &&
        request.status !== "Призупинена" &&
        request.status !== "Закрита"
      );


    default:

      return true;

  }

}


// ============================================================
// COUNTERS
// ============================================================

function updateCounters() {

  const all =
    STATE.requests.length;


  const newCount =
    STATE.requests.filter(
      item =>
        item.status === "Нова"
    ).length;


  const paused =
    STATE.requests.filter(
      item =>
        item.status === "Призупинена"
    ).length;


  const overdue =
    STATE.requests.filter(
      item =>
        item.slaOverdue === "Так" ||
        item.dateOverdue === "Так"
    ).length;


  const work =
    STATE.requests.filter(
      item =>
        item.status !== "Нова" &&
        item.status !== "Призупинена" &&
        item.status !== "Закрита"
    ).length;


  setText(
    "countAll",
    all
  );

  setText(
    "countNew",
    newCount
  );

  setText(
    "countWork",
    work
  );

  setText(
    "countPaused",
    paused
  );

  setText(
    "countOverdue",
    overdue
  );

  setText(
    "sidebarCount",
    all
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


  if (!STATE.executor) {

    body.innerHTML = `
      <tr>
        <td colspan="8">
          <div class="empty-state">
            Оберіть виконавця у верхній частині форми
          </div>
        </td>
      </tr>
    `;

    setText(
      "tableSummary",
      "Показано 0 заявок"
    );

    updateCounters();

    return;

  }


  if (
    STATE.filteredRequests.length === 0
  ) {

    body.innerHTML = `
      <tr>
        <td colspan="8">
          <div class="empty-state">
            Заявок за обраними умовами немає
          </div>
        </td>
      </tr>
    `;

    setText(
      "tableSummary",
      "Показано 0 заявок"
    );

    return;

  }


  body.innerHTML =
    STATE.filteredRequests
      .map(request =>
        buildRequestRow(
          request
        )
      )
      .join("");


  setText(
    "tableSummary",
    `Показано ${STATE.filteredRequests.length} з ${STATE.requests.length} заявок`
  );


  bindRowButtons();

}


// ============================================================
// BUILD ROW
// ============================================================

function buildRequestRow(
  request
) {

  const statusClass =
    getStatusClass(
      request.status
    );


  const slaHtml =
    request.slaOverdue === "Так"
      ? `<span class="sla-overdue">Прострочено</span>`
      : request.sla > 0
        ? `<span class="sla-ok">${request.sla} год</span>`
        : `<span class="subtext">—</span>`;


  return `
    <tr>

      <td>
        <div class="request-id">
          ${escapeHtml(request.id)}
        </div>

        ${
          request.sd
            ? `<div class="subtext">${escapeHtml(request.sd)}</div>`
            : ""
        }
      </td>

      <td>
        ${formatExcelDate(request.created)}
      </td>

      <td>
        <div
          class="description-cell"
          title="${escapeHtml(request.description)}"
        >
          <strong>
            ${escapeHtml(request.description || "Без опису")}
          </strong>
        </div>

        <div class="subtext">
          ${escapeHtml(request.customer)}
        </div>
      </td>

      <td>
        ${escapeHtml(request.city)}
      </td>

      <td>
        ${escapeHtml(request.category)}
      </td>

      <td>
        <span class="badge ${statusClass}">
          ${escapeHtml(request.status)}
        </span>
      </td>

      <td>
        ${slaHtml}
      </td>

      <td>

        <div class="action-group">

          <button
            class="icon-button"
            data-action="view"
            data-id="${escapeHtml(request.id)}"
            title="Переглянути"
          >
            ◉
          </button>

          <button
            class="icon-button"
            data-action="edit"
            data-id="${escapeHtml(request.id)}"
            title="Редагувати"
          >
            ✎
          </button>

          <button
            class="icon-button process"
            data-action="process"
            data-id="${escapeHtml(request.id)}"
            title="Опрацювати"
          >
            ▶
          </button>

        </div>

      </td>

    </tr>
  `;

}


// ============================================================
// ROW BUTTONS
// ============================================================

function bindRowButtons() {

  document
    .querySelectorAll(
      "[data-action]"
    )
    .forEach(button => {

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


          if (action === "view") {

            openDetails(
              request
            );

          }


          if (action === "process") {

            openStatusModal(
              request
            );

          }


          if (action === "edit") {

            openEditModal(
              request
            );

          }

        }
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


  const rows = [
    ["Номер SD", request.sd],
    ["Категорія", request.category],
    ["Опис", request.description],
    ["Місто", request.city],
    ["Адреса", request.address],
    ["Дата створення", formatExcelDate(request.created)],
    ["Сума витрат", formatMoney(request.amount)],
    ["Замовник", request.customer],
    ["Статус", request.status],
    ["Виконавець", request.executor],
    ["Попередній статус", request.previousStatus],
    ["Поточний статус з", formatExcelDateTime(request.statusSince)],
    ["SLA", request.sla ? request.sla + " год" : "—"],
    ["Прострочено SLA", request.slaOverdue || "Ні"],
    ["Планова дата", formatExcelDate(request.plannedDate)],
    ["Причина призупинення", request.pauseReason],
    ["Коментар", request.comment]
  ];


  document
    .getElementById(
      "detailsContent"
    )
    .innerHTML =
      rows
        .map(item => `
          <div class="details-label">
            ${escapeHtml(item[0])}
          </div>

          <div class="details-value">
            ${escapeHtml(String(item[1] || "—"))}
          </div>
        `)
        .join("");


  openModal(
    "detailsModal"
  );

}


// ============================================================
// STATUS
// ============================================================

function openStatusModal(
  request
) {

  STATE.selectedRequest =
    request;


  setText(
    "statusModalTitle",
    `Опрацювання ${request.id}`
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
    .value = "";


  document
    .getElementById(
      "statusComment"
    )
    .value = "";


  document
    .getElementById(
      "pauseReason"
    )
    .value = "";


  clearMessage(
    "statusMessage"
  );


  handlePauseReason();


  openModal(
    "statusModal"
  );

}


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


  block.style.display =
    status === "Призупинена"
      ? "block"
      : "none";

}


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
    newStatus === "Призупинена" &&
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
        createId("EVT"),

      requestKey: "",

      requestId:
        request.id,

      operation:
        "STATUS_CHANGE",

      actor:
        STATE.executor,

      source:
        "Виконавець",

      executor:
        STATE.executor,

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


    setTimeout(() => {

      closeModal(
        "statusModal"
      );

    }, 900);

  } catch (error) {

    showError(
      "statusMessage",
      getErrorText(error)
    );

  } finally {

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


  if (STATE.executor) {

    document
      .getElementById(
        "createExecutor"
      )
      .value =
        STATE.executor;

  }


  openModal(
    "createModal"
  );

}


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
        createId("EVT"),

      requestKey:
        createId("REQ"),

      requestId: "",

      operation:
        "CREATE",

      actor:
        STATE.executor ||
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


    setTimeout(() => {

      closeModal(
        "createModal"
      );

    }, 900);

  } catch (error) {

    showError(
      "createMessage",
      getErrorText(error)
    );

  } finally {

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

function openEditModal(
  request
) {

  STATE.selectedRequest =
    request;


  setText(
    "editTitle",
    `Редагування ${request.id}`
  );


  document.getElementById("editSD").value =
    request.sd;

  document.getElementById("editCategory").value =
    request.category;

  document.getElementById("editDescription").value =
    request.description;

  document.getElementById("editCity").value =
    request.city;

  document.getElementById("editAddress").value =
    request.address;

  document.getElementById("editCustomer").value =
    request.customer;

  document.getElementById("editAmount").value =
    request.amount || "";

  document.getElementById("editPlannedDate").value =
    excelSerialToInputDate(
      request.plannedDate
    );

  document.getElementById("editComment").value =
    request.comment;


  clearMessage(
    "editMessage"
  );


  openModal(
    "editModal"
  );

}


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
        createId("EVT"),

      requestKey: "",

      requestId:
        request.id,

      operation:
        "UPDATE",

      actor:
        STATE.executor,

      source:
        "Виконавець",

      executor:
        STATE.executor,

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


    setTimeout(() => {

      closeModal(
        "editModal"
      );

    }, 900);

  } catch (error) {

    showError(
      "editMessage",
      getErrorText(error)
    );

  } finally {

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

  await Excel.run(async context => {

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

  });

}


// ============================================================
// EMPLOYEE TABLE
// ============================================================

function getEmployeeTableName(
  employee
) {

  const map = {

    "Іванченко В.М.":
      "tbl_Ivanchenko",

    "Войцехівський Г.В.":
      "tbl_Voitsekhivskyi",

    "Ридванський П.С.":
      "tbl_Rydvanskyi",

    "Галько А.І.":
      "tbl_Halko",

    "Трунов Ю.О.":
      "tbl_Trunov",

    "Желясков Д.О.":
      "tbl_Zheliaskov",

    "Слепущенко О.О.":
      "tbl_Slepushchenko",

    "Сергеєв П.А.":
      "tbl_Serheiev",

    "Туровський В.О.":
      "tbl_Turovskyi"

  };


  const result =
    map[employee];


  if (!result) {

    throw new Error(
      "Не знайдено таблицю виконавця: " +
      employee
    );

  }


  return result;

}


// ============================================================
// STATUS COLOR
// ============================================================

function getStatusClass(
  status
) {

  if (
    status === "Нова"
  ) {
    return "status-new";
  }


  if (
    status === "Призупинена"
  ) {
    return "status-paused";
  }


  if (
    status === "Закрита"
  ) {
    return "status-closed";
  }


  if (
    status === "Виконання робіт" ||
    status === "Прийняття робіт"
  ) {
    return "status-execution";
  }


  if (
    status === "Погодження кошторису" ||
    status === "Погодження бюджету" ||
    status === "Укладання договору"
  ) {
    return "status-agreement";
  }


  return "status-work";

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
    document.getElementById(id);


  if (!select) {
    return;
  }


  select.innerHTML = "";


  const option =
    document.createElement(
      "option"
    );


  option.value = "";
  option.textContent =
    placeholder;


  select.appendChild(
    option
  );


  values.forEach(value => {

    const row =
      document.createElement(
        "option"
      );


    row.value =
      value;

    row.textContent =
      value;


    select.appendChild(
      row
    );

  });

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


// ============================================================

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
// CREATE CLEAR
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
    .forEach(id => {

      const element =
        document.getElementById(id);


      if (element) {
        element.value = "";
      }

    });


  clearMessage(
    "createMessage"
  );

}


// ============================================================
// LOADING
// ============================================================

function showLoading() {

  document
    .getElementById(
      "requestsBody"
    )
    .innerHTML = `
      <tr>
        <td colspan="8">
          <div class="empty-state">
            Завантаження заявок...
          </div>
        </td>
      </tr>
    `;

}


// ============================================================

function showTableError(
  text
) {

  document
    .getElementById(
      "requestsBody"
    )
    .innerHTML = `
      <tr>
        <td colspan="8">
          <div class="empty-state">
            Помилка: ${escapeHtml(text)}
          </div>
        </td>
      </tr>
    `;

}


// ============================================================
// VALUES
// ============================================================

function valueOf(
  id
) {

  return String(
    document
      .getElementById(id)
      ?.value ?? ""
  ).trim();

}


// ============================================================

function numberOf(
  id
) {

  const value =
    valueOf(id)
      .replace(",", ".");


  if (!value) {
    return 0;
  }


  const number =
    Number(value);


  return Number.isFinite(number)
    ? number
    : 0;

}


// ============================================================
// IDs
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
// LOCAL DATE
// ============================================================

function localTimestamp() {

  const date =
    new Date();


  const pad =
    number =>
      String(number)
        .padStart(
          2,
          "0"
        );


  return (
    date.getFullYear() +
    "-" +
    pad(date.getMonth() + 1) +
    "-" +
    pad(date.getDate()) +
    "T" +
    pad(date.getHours()) +
    ":" +
    pad(date.getMinutes()) +
    ":" +
    pad(date.getSeconds())
  );

}


// ============================================================
// EXCEL DATE DISPLAY
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
    typeof value === "number"
  ) {

    const date =
      new Date(
        Math.round(
          (value - 25569) *
          86400000
        )
      );


    return [
      String(
        date.getUTCDate()
      ).padStart(2, "0"),

      String(
        date.getUTCMonth() + 1
      ).padStart(2, "0"),

      date.getUTCFullYear()
    ].join(".");

  }


  return String(value);

}


// ============================================================

function formatExcelDateTime(
  value
) {

  if (
    typeof value !== "number"
  ) {

    return value
      ? String(value)
      : "—";

  }


  const date =
    new Date(
      Math.round(
        (value - 25569) *
        86400000
      )
    );


  return (
    String(
      date.getUTCDate()
    ).padStart(2, "0") +
    "." +

    String(
      date.getUTCMonth() + 1
    ).padStart(2, "0") +
    "." +

    date.getUTCFullYear() +
    " " +

    String(
      date.getUTCHours()
    ).padStart(2, "0") +
    ":" +

    String(
      date.getUTCMinutes()
    ).padStart(2, "0")
  );

}


// ============================================================

function excelSerialToInputDate(
  value
) {

  if (
    typeof value !== "number"
  ) {
    return "";
  }


  const date =
    new Date(
      Math.round(
        (value - 25569) *
        86400000
      )
    );


  return (
    date.getUTCFullYear() +
    "-" +

    String(
      date.getUTCMonth() + 1
    ).padStart(2, "0") +
    "-" +

    String(
      date.getUTCDate()
    ).padStart(2, "0")
  );

}


// ============================================================
// MONEY
// ============================================================

function formatMoney(
  value
) {

  const number =
    Number(value);


  if (
    !Number.isFinite(number) ||
    number === 0
  ) {
    return "—";
  }


  return number
    .toLocaleString(
      "uk-UA",
      {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      }
    ) + " грн";

}


// ============================================================
// TEXT
// ============================================================

function setText(
  id,
  value
) {

  const element =
    document.getElementById(id);


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
    document.getElementById(id);


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
  text
) {

  const element =
    document.getElementById(id);


  if (!element) {
    return;
  }


  element.className =
    "message success";

  element.textContent =
    "✓ " + text;

}


// ============================================================

function showError(
  id,
  text
) {

  const element =
    document.getElementById(id);


  if (!element) {
    return;
  }


  element.className =
    "message error";

  element.textContent =
    "Помилка: " + text;

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
    document.getElementById(id);


  if (!button) {
    return;
  }


  button.disabled =
    busy;

  button.textContent =
    caption;

}


// ============================================================
// ERROR
// ============================================================

function getErrorText(
  error
) {

  if (
    error &&
    error.message
  ) {
    return error.message;
  }


  return String(error);

}


// ============================================================
// ESCAPE
// ============================================================

function escapeHtml(
  value
) {

  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");

}
