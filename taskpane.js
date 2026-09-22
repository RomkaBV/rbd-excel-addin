const RBD = {

  source: "Керівник",

  actor: "Керівник"

};


let executorRequests = [];


// ============================================================
// START
// ============================================================

Office.onReady(async (info) => {

  initTabs();

  initUi();


  if (
    info.host !==
    Office.HostType.Excel
  ) {

    showError(
      "Надбудову потрібно відкрити в Excel."
    );

    return;

  }


  try {

    await loadDictionaries();

  }

  catch (error) {

    showError(
      "Не вдалося завантажити довідники: " +
      getErrorText(error)
    );

  }

});


// ============================================================
// TABS
// ============================================================

function initTabs() {

  const buttons =
    document.querySelectorAll(
      ".tab-button"
    );


  buttons.forEach(button => {

    button.addEventListener(
      "click",
      () => {

        const target =
          button.dataset.tab;


        document
          .querySelectorAll(
            ".tab-button"
          )
          .forEach(item => {

            item.classList.remove(
              "active"
            );

          });


        document
          .querySelectorAll(
            ".tab-section"
          )
          .forEach(section => {

            section.classList.remove(
              "active"
            );

          });


        button.classList.add(
          "active"
        );


        document
          .getElementById(target)
          ?.classList.add(
            "active"
          );

      }
    );

  });

}


// ============================================================
// UI EVENTS
// ============================================================

function initUi() {

  const description =
    document.getElementById(
      "description"
    );


  const comment =
    document.getElementById(
      "comment"
    );


  description?.addEventListener(
    "input",
    () => {

      const counter =
        document.getElementById(
          "descriptionCount"
        );


      if (counter) {

        counter.textContent =
          String(
            description.value.length
          );

      }

    }
  );


  comment?.addEventListener(
    "input",
    () => {

      const counter =
        document.getElementById(
          "commentCount"
        );


      if (counter) {

        counter.textContent =
          String(
            comment.value.length
          );

      }

    }
  );


  document
    .getElementById(
      "clearButton"
    )
    ?.addEventListener(
      "click",
      () => clearForm(true)
    );


  document
    .getElementById(
      "createButton"
    )
    ?.addEventListener(
      "click",
      createRequest
    );


  document
    .getElementById(
      "cabinetExecutor"
    )
    ?.addEventListener(
      "change",
      loadExecutorRequests
    );


  document
    .getElementById(
      "requestSelect"
    )
    ?.addEventListener(
      "change",
      loadSelectedRequest
    );


  document
    .getElementById(
      "newStatus"
    )
    ?.addEventListener(
      "change",
      handleNewStatusChange
    );


  document
    .getElementById(
      "changeStatusButton"
    )
    ?.addEventListener(
      "click",
      changeRequestStatus
    );


  document
    .getElementById(
      "refreshRequestsButton"
    )
    ?.addEventListener(
      "click",
      loadExecutorRequests
    );


  initFiles();

}


// ============================================================
// LOAD DICTIONARIES
// ============================================================

async function loadDictionaries() {

  await Excel.run(
    async context => {

      const categories =
        context.workbook.tables.getItem(
          "tbl_RBD_Categories"
        );


      const cities =
        context.workbook.tables.getItem(
          "tbl_RBD_Cities"
        );


      const employees =
        context.workbook.tables.getItem(
          "tbl_RBD_Employees"
        );


      const sla =
        context.workbook.tables.getItem(
          "tbl_SLA"
        );


      const categoryRange =
        categories.getDataBodyRange();


      const cityRange =
        cities.getDataBodyRange();


      const employeeRange =
        employees.getDataBodyRange();


      const slaRange =
        sla.getDataBodyRange();


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


      fillSelect(
        "category",
        categoryRange.values,
        "Оберіть категорію"
      );


      fillSelect(
        "city",
        cityRange.values,
        "Оберіть місто"
      );


      fillSelect(
        "executor",
        employeeRange.values,
        "Оберіть виконавця"
      );


      fillSelect(
        "cabinetExecutor",
        employeeRange.values,
        "Оберіть виконавця"
      );


      fillSelect(
        "newStatus",
        slaRange.values,
        "Оберіть новий статус"
      );

    }
  );

}


// ============================================================
// FILL SELECT
// ============================================================

function fillSelect(
  id,
  rows,
  placeholder
) {

  const select =
    document.getElementById(id);


  if (!select) {
    return;
  }


  select.innerHTML = "";


  const emptyOption =
    document.createElement(
      "option"
    );


  emptyOption.value = "";

  emptyOption.textContent =
    placeholder;


  select.appendChild(
    emptyOption
  );


  rows.forEach(row => {

    const value =
      String(
        row[0] ?? ""
      ).trim();


    if (!value) {
      return;
    }


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

  });

}


// ============================================================
// CREATE REQUEST
// ============================================================

async function createRequest() {

  clearMessage();


  const data = {

    sd:
      valueOf("sd"),

    category:
      valueOf("category"),

    description:
      valueOf("description"),

    comment:
      valueOf("comment"),

    city:
      valueOf("city"),

    address:
      valueOf("address"),

    customer:
      valueOf("customer"),

    executor:
      valueOf("executor"),

    amount:
      numberOf("amount"),

    plannedDate:
      valueOf("plannedDate")

  };


  const validation =
    validateCreate(
      data
    );


  if (validation) {

    showError(
      validation
    );

    return;

  }


  const button =
    document.getElementById(
      "createButton"
    );


  try {

    if (button) {

      button.disabled =
        true;


      button.textContent =
        "ПЕРЕДАЄМО В ЧЕРГУ...";

    }


    const eventId =
      createId(
        "EVT"
      );


    const requestKey =
      createId(
        "REQ"
      );


    await enqueue({

      eventId,

      requestKey,

      requestId: "",

      operation:
        "CREATE",

      actor:
        RBD.actor,

      source:
        RBD.source,

      executor:
        data.executor,

      expectedStatus: "",

      newStatus:
        "Нова",

      sd:
        data.sd,

      category:
        data.category,

      description:
        data.description,

      city:
        data.city,

      address:
        data.address,

      amount:
        data.amount,

      customer:
        data.customer,

      plannedDate:
        data.plannedDate,

      comment:
        data.comment,

      pauseReason: "",

      newExecutor: ""

    });


    showSuccess(
      "Заявку передано в чергу."
    );


    clearForm(
      false
    );

  }

  catch (error) {

    showError(
      "Не вдалося створити заявку: " +
      getErrorText(error)
    );

  }

  finally {

    if (button) {

      button.disabled =
        false;


      button.textContent =
        "СТВОРИТИ ЗАЯВКУ";

    }

  }

}


// ============================================================
// VALIDATE CREATE
// ============================================================

function validateCreate(
  data
) {

  if (!data.category) {

    return "Оберіть категорію.";

  }


  if (!data.description) {

    return "Заповніть опис потреби.";

  }


  if (!data.city) {

    return "Оберіть місто.";

  }


  if (!data.customer) {

    return "Заповніть замовника.";

  }


  if (!data.executor) {

    return "Оберіть виконавця.";

  }


  return "";

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


      /*
        ВАЖЛИВО:
        додаємо тільки один рядок у чергу.
        Основні таблиці JS не змінює.
      */

      queue.rows.add(
        null,
        [row]
      );


      await context.sync();

    }
  );

}


// ============================================================
// LOAD EXECUTOR REQUESTS
// ============================================================

async function loadExecutorRequests() {

  clearStatusMessage();


  executorRequests = [];


  const executor =
    valueOf(
      "cabinetExecutor"
    );


  const requestSelect =
    document.getElementById(
      "requestSelect"
    );


  clearSelectedRequest();


  if (!requestSelect) {
    return;
  }


  if (!executor) {

    requestSelect.innerHTML =
      `
      <option value="">
        Спочатку оберіть виконавця
      </option>
      `;


    return;

  }


  requestSelect.innerHTML =
    `
    <option value="">
      Завантаження...
    </option>
    `;


  try {

    const requests =
      await Excel.run(
        async context => {

          const tableName =
            getEmployeeTableName(
              executor
            );


          const table =
            context.workbook.tables.getItem(
              tableName
            );


          const headerRange =
            table.getHeaderRowRange();


          table.rows.load(
            "count"
          );


          headerRange.load(
            "values"
          );


          await context.sync();


          const headers =
            headerRange.values[0];


          /*
            Якщо таблиця порожня,
            не звертаємось до DataBodyRange.
          */

          if (
            table.rows.count === 0
          ) {

            return [];

          }


          const bodyRange =
            table.getDataBodyRange();


          bodyRange.load(
            "values"
          );


          await context.sync();


          const rows =
            bodyRange.values;


          const idIndex =
            findHeader(
              headers,
              "ID"
            );


          const descriptionIndex =
            findHeader(
              headers,
              "Опис"
            );


          const statusIndex =
            findHeader(
              headers,
              "Статус"
            );


          const cityIndex =
            findHeader(
              headers,
              "Місто"
            );


          const sdIndex =
            findHeader(
              headers,
              "Номер заявки SD"
            );


          const customerIndex =
            findHeader(
              headers,
              "Замовник"
            );


          const categoryIndex =
            findHeader(
              headers,
              "Категорія"
            );


          const plannedIndex =
            findHeader(
              headers,
              "Планова дата завершення"
            );


          const result = [];


          rows.forEach(row => {

            const id =
              String(
                row[idIndex] ??
                ""
              ).trim();


            if (!id) {
              return;
            }


            result.push({

              id,

              description:
                getStringValue(
                  row,
                  descriptionIndex
                ),

              status:
                getStringValue(
                  row,
                  statusIndex
                ),

              city:
                getStringValue(
                  row,
                  cityIndex
                ),

              sd:
                getStringValue(
                  row,
                  sdIndex
                ),

              customer:
                getStringValue(
                  row,
                  customerIndex
                ),

              category:
                getStringValue(
                  row,
                  categoryIndex
                ),

              plannedDate:
                plannedIndex >= 0
                  ? row[
                      plannedIndex
                    ]
                  : ""

            });

          });


          return result;

        }
      );


    executorRequests =
      requests;


    requestSelect.innerHTML =
      `
      <option value="">
        Оберіть заявку
      </option>
      `;


    if (
      executorRequests.length === 0
    ) {

      requestSelect.innerHTML =
        `
        <option value="">
          Активних заявок немає
        </option>
        `;


      return;

    }


    executorRequests.forEach(
      request => {

        const option =
          document.createElement(
            "option"
          );


        option.value =
          request.id;


        option.textContent =
          buildRequestCaption(
            request
          );


        requestSelect.appendChild(
          option
        );

      }
    );

  }

  catch (error) {

    requestSelect.innerHTML =
      `
      <option value="">
        Помилка завантаження
      </option>
      `;


    showStatusError(
      getErrorText(error)
    );

  }

}


// ============================================================
// REQUEST CAPTION
// ============================================================

function buildRequestCaption(
  request
) {

  let caption =
    request.id;


  if (request.city) {

    caption +=
      " | " +
      request.city;

  }


  if (request.description) {

    let description =
      request.description;


    if (
      description.length > 60
    ) {

      description =
        description.substring(
          0,
          60
        ) + "...";

    }


    caption +=
      " | " +
      description;

  }


  return caption;

}


// ============================================================
// SELECT REQUEST
// ============================================================

function loadSelectedRequest() {

  clearStatusMessage();


  const requestId =
    valueOf(
      "requestSelect"
    );


  clearSelectedRequest();


  if (!requestId) {
    return;
  }


  const request =
    executorRequests.find(
      item =>
        item.id ===
        requestId
    );


  if (!request) {

    showStatusError(
      "Заявку не знайдено."
    );

    return;

  }


  const currentStatus =
    document.getElementById(
      "currentStatus"
    );


  if (currentStatus) {

    currentStatus.value =
      request.status;

  }


  showRequestInfo(
    request
  );

}


// ============================================================
// REQUEST INFO
// ============================================================

function showRequestInfo(
  request
) {

  const container =
    document.getElementById(
      "requestInfo"
    );


  const title =
    document.getElementById(
      "requestInfoTitle"
    );


  const text =
    document.getElementById(
      "requestInfoText"
    );


  if (
    !container ||
    !title ||
    !text
  ) {

    return;

  }


  title.textContent =
    request.id +
    (
      request.sd
        ? " | SD: " +
          request.sd
        : ""
    );


  const parts = [];


  if (request.category) {

    parts.push(
      "Категорія: " +
      request.category
    );

  }


  if (request.city) {

    parts.push(
      "Місто: " +
      request.city
    );

  }


  if (request.customer) {

    parts.push(
      "Замовник: " +
      request.customer
    );

  }


  if (request.description) {

    parts.push(
      "Опис: " +
      request.description
    );

  }


  text.textContent =
    parts.join(
      " • "
    );


  container.classList.add(
    "visible"
  );

}


// ============================================================
// CLEAR SELECTED REQUEST
// ============================================================

function clearSelectedRequest() {

  const currentStatus =
    document.getElementById(
      "currentStatus"
    );


  const newStatus =
    document.getElementById(
      "newStatus"
    );


  const pauseReason =
    document.getElementById(
      "pauseReason"
    );


  const statusComment =
    document.getElementById(
      "statusComment"
    );


  const requestInfo =
    document.getElementById(
      "requestInfo"
    );


  if (currentStatus) {

    currentStatus.value = "";

  }


  if (newStatus) {

    newStatus.value = "";

  }


  if (pauseReason) {

    pauseReason.value = "";

  }


  if (statusComment) {

    statusComment.value = "";

  }


  if (requestInfo) {

    requestInfo.classList.remove(
      "visible"
    );

  }


  handleNewStatusChange();

}


// ============================================================
// STATUS SELECT
// ============================================================

function handleNewStatusChange() {

  const newStatus =
    valueOf(
      "newStatus"
    );


  const block =
    document.getElementById(
      "pauseReasonBlock"
    );


  const reason =
    document.getElementById(
      "pauseReason"
    );


  if (!block) {
    return;
  }


  if (
    newStatus ===
    "Призупинена"
  ) {

    block.style.display =
      "block";

  }

  else {

    block.style.display =
      "none";


    if (reason) {

      reason.value =
        "";

    }

  }

}


// ============================================================
// CHANGE STATUS
// ============================================================

async function changeRequestStatus() {

  clearStatusMessage();


  const executor =
    valueOf(
      "cabinetExecutor"
    );


  const requestId =
    valueOf(
      "requestSelect"
    );


  const currentStatus =
    valueOf(
      "currentStatus"
    );


  const newStatus =
    valueOf(
      "newStatus"
    );


  const comment =
    valueOf(
      "statusComment"
    );


  const pauseReason =
    valueOf(
      "pauseReason"
    );


  if (!executor) {

    showStatusError(
      "Оберіть виконавця."
    );

    return;

  }


  if (!requestId) {

    showStatusError(
      "Оберіть заявку."
    );

    return;

  }


  if (!currentStatus) {

    showStatusError(
      "Не визначено поточний статус."
    );

    return;

  }


  if (!newStatus) {

    showStatusError(
      "Оберіть новий статус."
    );

    return;

  }


  if (
    currentStatus ===
    newStatus
  ) {

    showStatusError(
      "Новий статус збігається з поточним."
    );

    return;

  }


  if (
    newStatus ===
      "Призупинена" &&
    !pauseReason
  ) {

    showStatusError(
      "Для призупинення обов'язково вкажіть причину."
    );

    return;

  }


  const button =
    document.getElementById(
      "changeStatusButton"
    );


  try {

    if (button) {

      button.disabled =
        true;


      button.textContent =
        "ПЕРЕДАЄМО В ЧЕРГУ...";

    }


    await enqueue({

      eventId:
        createId(
          "EVT"
        ),

      requestKey: "",

      requestId,

      operation:
        "STATUS_CHANGE",

      actor:
        executor,

      source:
        "Виконавець",

      executor,

      expectedStatus:
        currentStatus,

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


    showStatusSuccess(

      "Зміну статусу передано в чергу: " +
      currentStatus +
      " → " +
      newStatus +
      ". Після запуску 03_ProcessQueue натисніть «Оновити заявки»."

    );


    const newStatusField =
      document.getElementById(
        "newStatus"
      );


    const commentField =
      document.getElementById(
        "statusComment"
      );


    const pauseField =
      document.getElementById(
        "pauseReason"
      );


    if (newStatusField) {

      newStatusField.value =
        "";

    }


    if (commentField) {

      commentField.value =
        "";

    }


    if (pauseField) {

      pauseField.value =
        "";

    }


    handleNewStatusChange();

  }

  catch (error) {

    showStatusError(
      "Не вдалося передати зміну статусу: " +
      getErrorText(error)
    );

  }

  finally {

    if (button) {

      button.disabled =
        false;


      button.textContent =
        "ЗМІНИТИ СТАТУС";

    }

  }

}


// ============================================================
// EMPLOYEE → TABLE
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


  const tableName =
    map[employee];


  if (!tableName) {

    throw new Error(
      "Не визначена таблиця для виконавця: " +
      employee
    );

  }


  return tableName;

}


// ============================================================
// FILES UI
// ============================================================

function initFiles() {

  const fileInput =
    document.getElementById(
      "fileInput"
    );


  const fileButton =
    document.getElementById(
      "fileButton"
    );


  const fileBar =
    document.getElementById(
      "fileBar"
    );


  const fileText =
    document.getElementById(
      "fileText"
    );


  if (
    !fileInput ||
    !fileButton ||
    !fileBar ||
    !fileText
  ) {

    return;

  }


  fileButton.addEventListener(
    "click",
    () => {

      fileInput.click();

    }
  );


  fileInput.addEventListener(
    "change",
    () => {

      showSelectedFiles(
        fileInput.files
      );

    }
  );


  fileBar.addEventListener(
    "dragover",
    event => {

      event.preventDefault();


      fileBar.classList.add(
        "dragging"
      );

    }
  );


  fileBar.addEventListener(
    "dragleave",
    () => {

      fileBar.classList.remove(
        "dragging"
      );

    }
  );


  fileBar.addEventListener(
    "drop",
    event => {

      event.preventDefault();


      fileBar.classList.remove(
        "dragging"
      );


      showSelectedFiles(
        event.dataTransfer.files
      );

    }
  );

}


// ============================================================
// SHOW FILES
// ============================================================

function showSelectedFiles(
  files
) {

  const fileText =
    document.getElementById(
      "fileText"
    );


  if (!fileText) {
    return;
  }


  if (
    !files ||
    files.length === 0
  ) {

    fileText.textContent =
      "Файли не вибрані";


    return;

  }


  if (
    files.length === 1
  ) {

    fileText.textContent =
      files[0].name;


    return;

  }


  fileText.textContent =
    `${files.length} файлів вибрано`;

}


// ============================================================
// CLEAR CREATE FORM
// ============================================================

function clearForm(
  clearMessageToo = true
) {

  const ids = [

    "sd",

    "category",

    "description",

    "comment",

    "city",

    "address",

    "customer",

    "executor",

    "amount",

    "plannedDate"

  ];


  ids.forEach(id => {

    const element =
      document.getElementById(id);


    if (element) {

      element.value =
        "";

    }

  });


  const descriptionCount =
    document.getElementById(
      "descriptionCount"
    );


  const commentCount =
    document.getElementById(
      "commentCount"
    );


  const fileText =
    document.getElementById(
      "fileText"
    );


  const fileInput =
    document.getElementById(
      "fileInput"
    );


  if (descriptionCount) {

    descriptionCount.textContent =
      "0";

  }


  if (commentCount) {

    commentCount.textContent =
      "0";

  }


  if (fileText) {

    fileText.textContent =
      "Файли не вибрані";

  }


  if (fileInput) {

    fileInput.value =
      "";

  }


  if (clearMessageToo) {

    clearMessage();

  }

}


// ============================================================
// VALUE HELPERS
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

  const raw =
    valueOf(id)
      .replace(
        ",",
        "."
      );


  if (!raw) {

    return 0;

  }


  const result =
    Number(raw);


  return Number.isFinite(
    result
  )
    ? result
    : 0;

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


  const pad =
    value =>
      String(value)
        .padStart(
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
// HEADER
// ============================================================

function findHeader(
  headers,
  name
) {

  return headers.indexOf(
    name
  );

}


// ============================================================
// SAFE STRING
// ============================================================

function getStringValue(
  row,
  index
) {

  if (
    index < 0
  ) {

    return "";

  }


  return String(
    row[index] ?? ""
  ).trim();

}


// ============================================================
// CREATE MESSAGE
// ============================================================

function clearMessage() {

  const message =
    document.getElementById(
      "message"
    );


  if (!message) {
    return;
  }


  message.className =
    "message";


  message.textContent =
    "";

}


// ============================================================

function showSuccess(
  text
) {

  const message =
    document.getElementById(
      "message"
    );


  if (!message) {
    return;
  }


  message.className =
    "message success";


  message.textContent =
    "✓ " + text;

}


// ============================================================

function showError(
  text
) {

  const message =
    document.getElementById(
      "message"
    );


  if (!message) {
    return;
  }


  message.className =
    "message error";


  message.textContent =
    "Помилка: " + text;

}


// ============================================================
// STATUS MESSAGE
// ============================================================

function clearStatusMessage() {

  const message =
    document.getElementById(
      "statusMessage"
    );


  if (!message) {
    return;
  }


  message.className =
    "message";


  message.textContent =
    "";

}


// ============================================================

function showStatusSuccess(
  text
) {

  const message =
    document.getElementById(
      "statusMessage"
    );


  if (!message) {
    return;
  }


  message.className =
    "message success";


  message.textContent =
    "✓ " + text;

}


// ============================================================

function showStatusError(
  text
) {

  const message =
    document.getElementById(
      "statusMessage"
    );


  if (!message) {
    return;
  }


  message.className =
    "message error";


  message.textContent =
    "Помилка: " + text;

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
