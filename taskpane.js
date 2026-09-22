const RBD = {

  source: "Керівник",

  actor: "Керівник"

};


// ============================================================
// START
// ============================================================

Office.onReady(async (info) => {

  initUi();

  if (info.host !== Office.HostType.Excel) {

    console.log(
      "Форма відкрита поза Excel."
    );

    return;
  }


  try {

    await loadDictionaries();

  } catch (error) {

    showError(
      "Не вдалося завантажити довідники: " +
      getErrorText(error)
    );

  }

});


// ============================================================
// UI
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
    .getElementById("clearButton")
    ?.addEventListener(
      "click",
      clearForm
    );


  document
    .getElementById("createButton")
    ?.addEventListener(
      "click",
      createRequest
    );


  initFiles();

}


// ============================================================
// ДОВІДНИКИ З EXCEL
// ============================================================

async function loadDictionaries() {

  await Excel.run(async context => {

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


    const categoryRange =
      categories.getRangeBetweenHeaderAndTotal();

    const cityRange =
      cities.getRangeBetweenHeaderAndTotal();

    const employeeRange =
      employees.getRangeBetweenHeaderAndTotal();


    categoryRange.load("values");
    cityRange.load("values");
    employeeRange.load("values");


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

  });

}


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


  const empty =
    document.createElement(
      "option"
    );

  empty.value = "";
  empty.textContent =
    placeholder;

  select.appendChild(
    empty
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
// CREATE
// ============================================================

async function createRequest() {

  clearMessage();


  const data = {

    sd: valueOf("sd"),

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
    validateCreate(data);


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


  if (button) {

    button.disabled =
      true;

    button.textContent =
      "ПЕРЕДАЄМО В ЧЕРГУ...";

  }


  try {

    const eventId =
      createId("EVT");

    const requestKey =
      createId("REQ");


    await enqueue({

      eventId,

      requestKey,

      requestId: "",

      operation: "CREATE",

      actor:
        RBD.actor,

      source:
        RBD.source,

      executor:
        data.executor,

      expectedStatus: "",

      newStatus: "Нова",

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
      "Заявку передано в чергу. Event ID: " +
      eventId
    );


    clearForm(
      false
    );


  } catch (error) {

    showError(
      "Не вдалося передати заявку в чергу: " +
      getErrorText(error)
    );

  } finally {

    if (button) {

      button.disabled =
        false;

      button.innerHTML =
        "➤ &nbsp; СТВОРИТИ ЗАЯВКУ";

    }

  }

}


// ============================================================
// УНІВЕРСАЛЬНИЙ ЗАПИС У ЧЕРГУ
// ============================================================

async function enqueue(event) {

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
      [row],
      true
    );


    await context.sync();

  });

}


// ============================================================
// STATUS CHANGE
// Використаємо пізніше в кабінеті виконавця
// ============================================================

async function enqueueStatusChange(
  requestId,
  executor,
  currentStatus,
  newStatus,
  comment = "",
  pauseReason = ""
) {

  if (
    newStatus === "Призупинена" &&
    !pauseReason.trim()
  ) {

    throw new Error(
      "Для статусу «Призупинена» потрібна причина."
    );

  }


  await enqueue({

    eventId:
      createId("EVT"),

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

}


// ============================================================
// UPDATE
// ============================================================

async function enqueueUpdate(
  requestId,
  executor,
  changes
) {

  await enqueue({

    eventId:
      createId("EVT"),

    requestKey: "",

    requestId,

    operation:
      "UPDATE",

    actor:
      executor,

    source:
      "Виконавець",

    executor,

    expectedStatus: "",

    newStatus: "",

    sd:
      changes.sd ?? "",

    category:
      changes.category ?? "",

    description:
      changes.description ?? "",

    city:
      changes.city ?? "",

    address:
      changes.address ?? "",

    amount:
      changes.amount ?? "",

    customer:
      changes.customer ?? "",

    plannedDate:
      changes.plannedDate ?? "",

    comment:
      changes.comment ?? "",

    pauseReason: "",

    newExecutor: ""

  });

}


// ============================================================
// ASSIGN
// ============================================================

async function enqueueAssign(
  requestId,
  currentExecutor,
  newExecutor,
  comment = ""
) {

  await enqueue({

    eventId:
      createId("EVT"),

    requestKey: "",

    requestId,

    operation:
      "ASSIGN",

    actor:
      currentExecutor,

    source:
      "Виконавець",

    executor:
      currentExecutor,

    expectedStatus: "",

    newStatus: "",

    sd: "",
    category: "",
    description: "",
    city: "",
    address: "",
    amount: "",
    customer: "",
    plannedDate: "",

    comment,

    pauseReason: "",

    newExecutor

  });

}


// ============================================================
// ЗРОБИМО ДОСТУПНИМ ДЛЯ МАЙБУТНЬОГО UI
// ============================================================

window.RBDQueue = {

  statusChange:
    enqueueStatusChange,

  update:
    enqueueUpdate,

  assign:
    enqueueAssign

};


// ============================================================
// VALIDATION
// ============================================================

function validateCreate(data) {

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
// HELPERS
// ============================================================

function valueOf(id) {

  return String(
    document
      .getElementById(id)
      ?.value ?? ""
  ).trim();

}


function numberOf(id) {

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


function createId(prefix) {

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


function localTimestamp() {

  const d =
    new Date();


  const pad =
    value =>
      String(value)
        .padStart(2, "0");


  return (

    d.getFullYear() +
    "-" +

    pad(
      d.getMonth() + 1
    ) +
    "-" +

    pad(
      d.getDate()
    ) +
    "T" +

    pad(
      d.getHours()
    ) +
    ":" +

    pad(
      d.getMinutes()
    ) +
    ":" +

    pad(
      d.getSeconds()
    )

  );

}


// ============================================================
// FILES — ПОКИ ЛИШЕ UI
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
    () => fileInput.click()
  );


  fileInput.addEventListener(
    "change",
    () => showFiles(
      fileInput.files
    )
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
    () =>
      fileBar.classList.remove(
        "dragging"
      )
  );


  fileBar.addEventListener(
    "drop",
    event => {

      event.preventDefault();

      fileBar.classList.remove(
        "dragging"
      );

      showFiles(
        event.dataTransfer.files
      );

    }
  );


  function showFiles(files) {

    if (
      !files ||
      files.length === 0
    ) {

      fileText.textContent =
        "Файли не вибрані";

      return;

    }


    if (files.length === 1) {

      fileText.textContent =
        files[0].name;

      return;

    }


    fileText.textContent =
      `${files.length} файлів вибрано`;

  }

}


// ============================================================
// CLEAR / MESSAGE
// ============================================================

function clearForm(
  clearMessageToo = true
) {

  [
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

  ].forEach(id => {

    const el =
      document.getElementById(id);

    if (el) {
      el.value = "";
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
    descriptionCount.textContent = "0";
  }

  if (commentCount) {
    commentCount.textContent = "0";
  }

  if (fileText) {
    fileText.textContent =
      "Файли не вибрані";
  }

  if (fileInput) {
    fileInput.value = "";
  }


  if (clearMessageToo) {
    clearMessage();
  }

}


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


function showSuccess(text) {

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


function showError(text) {

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


function getErrorText(error) {

  if (
    error &&
    typeof error.message ===
      "string"
  ) {

    return error.message;

  }

  return String(error);

}
