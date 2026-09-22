Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    return;
  }

  document
    .getElementById("createButton")
    .addEventListener("click", createRequest);

  try {
    await loadDictionaries();
    await refreshDashboard();
  } catch (error) {
    showError("Помилка завантаження: " + getErrorMessage(error));
  }
});


// ============================================================
// ДОВІДНИКИ
// ============================================================

async function loadDictionaries() {

  await Excel.run(async (context) => {

    const workbook = context.workbook;

    const categories =
      workbook.names
        .getItem("RBD_Categories")
        .getRange();

    const cities =
      workbook.names
        .getItem("RBD_Cities")
        .getRange();

    const employees =
      workbook.names
        .getItem("RBD_Employees")
        .getRange();

    categories.load("values");
    cities.load("values");
    employees.load("values");

    await context.sync();

    fillSelect(
      "category",
      categories.values,
      "Оберіть категорію..."
    );

    fillSelect(
      "city",
      cities.values,
      "Оберіть місто..."
    );

    fillSelect(
      "executor",
      employees.values,
      "Оберіть виконавця..."
    );
  });
}


function fillSelect(id, values, placeholder) {

  const select =
    document.getElementById(id);

  select.innerHTML = "";

  const firstOption =
    document.createElement("option");

  firstOption.value = "";
  firstOption.textContent = placeholder;

  select.appendChild(firstOption);

  values.forEach((row) => {

    const value =
      String(row[0] || "").trim();

    if (!value) {
      return;
    }

    const option =
      document.createElement("option");

    option.value = value;
    option.textContent = value;

    select.appendChild(option);
  });
}


// ============================================================
// СТВОРЕННЯ ЗАЯВКИ
// ============================================================

async function createRequest() {

  const button =
    document.getElementById("createButton");

  clearMessage();

  const data = {
    sd: getValue("sd"),
    category: getValue("category"),
    description: getValue("description"),
    city: getValue("city"),
    address: getValue("address"),
    amount: parseAmount(getValue("amount")),
    customer: getValue("customer"),
    executor: getValue("executor"),
    plannedDate: getValue("plannedDate"),
    comment: getValue("comment")
  };


  if (!data.category) {
    showError("Оберіть категорію.");
    return;
  }

  if (!data.description) {
    showError("Заповніть опис.");
    return;
  }

  if (!data.city) {
    showError("Оберіть місто.");
    return;
  }

  if (!data.customer) {
    showError("Заповніть замовника.");
    return;
  }

  if (!data.executor) {
    showError("Оберіть виконавця.");
    return;
  }


  button.disabled = true;
  button.textContent = "СТВОРЕННЯ...";

  try {

    const newId =
      await saveRequest(data);

    clearForm();

    showSuccess(
      `Заявку ${newId} успішно створено`
    );

    await refreshDashboard();

  } catch (error) {

    showError(
      getErrorMessage(error)
    );

  } finally {

    button.disabled = false;
    button.textContent = "СТВОРИТИ ЗАЯВКУ";
  }
}


// ============================================================
// ЗАПИС У EXCEL
// ============================================================

async function saveRequest(data) {

  return Excel.run(async (context) => {

    const workbook =
      context.workbook;


    // ----------------------------------------------------------
    // ВИКОНАВЦІ
    // ----------------------------------------------------------

    const employeeTables = {

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


    const employeeTableName =
      employeeTables[data.executor];

    if (!employeeTableName) {
      throw new Error(
        "Для виконавця не визначена таблиця."
      );
    }


    // ----------------------------------------------------------
    // ТАБЛИЦІ
    // ----------------------------------------------------------

    const employeeTable =
      workbook.tables.getItem(
        employeeTableName
      );

    const baseTable =
      workbook.tables.getItem(
        "tbl_RBD_Base"
      );

    const registryTable =
      workbook.tables.getItem(
        "tbl_RBD_Registry"
      );

    const historyTable =
      workbook.tables.getItem(
        "tbl_StatusHistory"
      );

    const managerTable =
      workbook.tables.getItem(
        "tbl_ManagerRequests"
      );

    const slaTable =
      workbook.tables.getItem(
        "tbl_SLA"
      );


    // ----------------------------------------------------------
    // ЧИТАЄМО РЕЄСТР І SLA
    // ----------------------------------------------------------

    const registryRange =
      registryTable.getRange();

    const slaRange =
      slaTable.getRange();

    registryRange.load("values");
    slaRange.load("values");

    await context.sync();


    // ----------------------------------------------------------
    // НОВИЙ ID
    // ----------------------------------------------------------

    let maxId = 0;

    const registryValues =
      registryRange.values;

    for (
      let i = 1;
      i < registryValues.length;
      i++
    ) {

      const id =
        String(
          registryValues[i][0] || ""
        ).trim();

      if (!id.startsWith("RBD-")) {
        continue;
      }

      const number =
        Number(
          id.replace("RBD-", "")
        );

      if (
        !isNaN(number) &&
        number > maxId
      ) {
        maxId = number;
      }
    }


    const newId =
      "RBD-" +
      String(maxId + 1)
        .padStart(6, "0");


    // ----------------------------------------------------------
    // SLA НОВА
    // ----------------------------------------------------------

    let slaHours = 0;

    const slaValues =
      slaRange.values;

    for (
      let i = 1;
      i < slaValues.length;
      i++
    ) {

      const status =
        String(
          slaValues[i][0] || ""
        ).trim();

      if (status === "Нова") {

        slaHours =
          Number(
            slaValues[i][1]
          ) || 0;

        break;
      }
    }


    // ----------------------------------------------------------
    // ДАТИ
    // ----------------------------------------------------------

    const now =
      new Date();

    const createdDate =
      excelSerial(now);

    let plannedDate = "";

    if (data.plannedDate) {

      const date =
        new Date(
          data.plannedDate +
          "T00:00:00"
        );

      plannedDate =
        excelSerial(date);
    }


    // ----------------------------------------------------------
    // РЯДОК ЗАЯВКИ
    // ----------------------------------------------------------

    const requestRow = [

      data.sd,            // A
      data.category,      // B
      data.description,   // C
      data.city,          // D
      data.address,       // E

      createdDate,        // F

      data.amount,        // G
      data.customer,      // H

      "Нова",             // I

      data.executor,      // J

      newId,              // K

      "Керівник",         // L

      createdDate,        // M

      "Нова",             // N

      slaHours,           // O

      0,                  // P

      "Ні",               // Q

      plannedDate,        // R

      "Ні",               // S

      "",                 // T

      data.comment,       // U

      ""                  // V
    ];


    // ----------------------------------------------------------
    // ВИКОНАВЕЦЬ
    // ----------------------------------------------------------

    employeeTable.rows.add(
      null,
      [requestRow]
    );


    // ----------------------------------------------------------
    // ЗАГАЛЬНА БАЗА
    // ----------------------------------------------------------

    baseTable.rows.add(
      null,
      [requestRow]
    );


    // ----------------------------------------------------------
    // РЕЄСТР
    // ----------------------------------------------------------

    registryTable.rows.add(
      null,
      [[
        newId,
        data.sd,
        data.executor,
        "Керівник",
        "ACTIVE",
        createdDate,
        createdDate
      ]]
    );


    // ----------------------------------------------------------
    // ІСТОРІЯ
    // ----------------------------------------------------------

    historyTable.rows.add(
      null,
      [[
        `H-${newId}-001`,
        newId,
        data.sd,
        data.executor,
        "",
        "Нова",
        createdDate,
        "",
        0,
        slaHours,
        "Ні",
        data.comment,
        "Система"
      ]]
    );


    // ----------------------------------------------------------
    // КАБІНЕТ КЕРІВНИКА
    // ----------------------------------------------------------

    managerTable.rows.add(
      0,
      [[
        newId,
        createdDate,
        data.sd,
        data.category,
        data.city,
        data.address,
        data.customer,
        data.executor,
        "Нова",
        data.amount,
        plannedDate
      ]]
    );


    await context.sync();

    return newId;
  });
}


// ============================================================
// KPI + ОСТАННІ ЗАЯВКИ
// ============================================================

async function refreshDashboard() {

  await Excel.run(async (context) => {

    const workbook =
      context.workbook;

    const table =
      workbook.tables.getItem(
        "tbl_ManagerRequests"
      );

    const range =
      table.getRange();

    range.load("values");

    await context.sync();

    const allValues =
      range.values;

    if (allValues.length <= 1) {

      updateKpiElements(
        0,
        0,
        0,
        0
      );

      renderRequests([]);

      return;
    }


    const values =
      allValues.slice(1);


    let todayCount = 0;
    let monthCount = 0;
    let activeCount = 0;
    let closedCount = 0;


    const now =
      new Date();

    const todaySerial =
      Math.floor(
        excelSerial(now)
      );


    values.forEach((row) => {

      const id =
        String(
          row[0] || ""
        ).trim();

      if (!id) {
        return;
      }


      const createdSerial =
        Number(row[1]);

      const status =
        String(
          row[8] || ""
        ).trim();


      if (
        !isNaN(createdSerial) &&
        createdSerial > 0
      ) {

        if (
          Math.floor(createdSerial) ===
          todaySerial
        ) {
          todayCount++;
        }


        const createdDate =
          excelDateFromSerial(
            createdSerial
          );


        if (
          createdDate.getFullYear() ===
            now.getFullYear() &&
          createdDate.getMonth() ===
            now.getMonth()
        ) {
          monthCount++;
        }
      }


      if (status === "Закрита") {
        closedCount++;
      } else {
        activeCount++;
      }
    });


    updateKpiElements(
      todayCount,
      monthCount,
      activeCount,
      closedCount
    );

    renderRequests(values);
  });
}


// ============================================================
// KPI DOM
// ============================================================

function updateKpiElements(
  today,
  month,
  active,
  closed
) {

  document
    .getElementById("kpiToday")
    .textContent = today;

  document
    .getElementById("kpiMonth")
    .textContent = month;

  document
    .getElementById("kpiActive")
    .textContent = active;

  document
    .getElementById("kpiClosed")
    .textContent = closed;
}


// ============================================================
// ОСТАННІ ЗАЯВКИ
// ============================================================

function renderRequests(values) {

  const body =
    document.getElementById(
      "requestsBody"
    );

  body.innerHTML = "";


  const rows =
    values
      .filter((row) => {
        return String(
          row[0] || ""
        ).trim();
      })
      .sort((a, b) => {
        return Number(b[1]) -
               Number(a[1]);
      })
      .slice(0, 20);


  if (rows.length === 0) {

    const tr =
      document.createElement("tr");

    const td =
      document.createElement("td");

    td.colSpan = 6;

    td.textContent =
      "Створених заявок поки немає.";

    tr.appendChild(td);

    body.appendChild(tr);

    return;
  }


  rows.forEach((row) => {

    const tr =
      document.createElement("tr");

    const createdDate =
      formatExcelDate(
        Number(row[1])
      );

    const cells = [
      row[0],
      createdDate,
      row[3],
      row[4],
      row[7],
      row[8]
    ];


    cells.forEach(
      (value, index) => {

        const td =
          document.createElement("td");

        td.textContent =
          String(value ?? "");

        if (index === 5) {
          td.className = "status";
        }

        tr.appendChild(td);
      }
    );


    body.appendChild(tr);
  });
}


// ============================================================
// ФОРМА
// ============================================================

function getValue(id) {

  const element =
    document.getElementById(id);

  if (!element) {
    return "";
  }

  return String(
    element.value || ""
  ).trim();
}


function parseAmount(value) {

  if (!value) {
    return 0;
  }

  const normalized =
    String(value)
      .replace(",", ".");

  const number =
    Number(normalized);

  if (isNaN(number)) {
    return 0;
  }

  return number;
}


function clearForm() {

  const ids = [
    "sd",
    "category",
    "description",
    "city",
    "address",
    "amount",
    "customer",
    "executor",
    "plannedDate",
    "comment"
  ];

  ids.forEach((id) => {

    const element =
      document.getElementById(id);

    if (element) {
      element.value = "";
    }
  });
}


// ============================================================
// ПОВІДОМЛЕННЯ
// ============================================================

function clearMessage() {

  const message =
    document.getElementById(
      "message"
    );

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

  message.className =
    "message message-success";

  message.textContent =
    "✓ " + text;
}


function showError(text) {

  const message =
    document.getElementById(
      "message"
    );

  message.className =
    "message message-error";

  message.textContent =
    "Помилка: " + text;
}


function getErrorMessage(error) {

  if (
    error &&
    typeof error.message === "string"
  ) {
    return error.message;
  }

  return String(error);
}


// ============================================================
// EXCEL DATE
// ============================================================

function excelSerial(date) {

  return (
    date.getTime() /
    86400000
  ) + 25569;
}


function excelDateFromSerial(serial) {

  return new Date(
    (serial - 25569) *
    86400000
  );
}


function formatExcelDate(serial) {

  if (
    !serial ||
    isNaN(serial)
  ) {
    return "";
  }

  const date =
    excelDateFromSerial(serial);

  return date.toLocaleString(
    "uk-UA",
    {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    }
  );
}
