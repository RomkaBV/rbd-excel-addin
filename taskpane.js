Office.onReady(async (info) => {

  // Якщо форма відкрита просто в браузері,
  // залишаємо значення, які є в HTML.
  if (info.host !== Office.HostType.Excel) {
    console.log("Форма відкрита поза Excel.");
    return;
  }

  try {

    await loadDictionaries();

    console.log(
      "Довідники з Excel успішно завантажені."
    );

  } catch (error) {

    console.error(
      "Помилка завантаження довідників:",
      error
    );

  }

});


// ============================================================
// ЗАВАНТАЖЕННЯ ДОВІДНИКІВ
// ============================================================

async function loadDictionaries() {

  await Excel.run(async (context) => {

    const workbook =
      context.workbook;


    // ----------------------------------------------------------
    // ІМЕНОВАНІ ДІАПАЗОНИ
    // ----------------------------------------------------------

    const categoriesName =
      workbook.names.getItemOrNullObject(
        "RBD_Categories"
      );

    const citiesName =
      workbook.names.getItemOrNullObject(
        "RBD_Cities"
      );

    const employeesName =
      workbook.names.getItemOrNullObject(
        "RBD_Employees"
      );


    categoriesName.load("isNullObject");
    citiesName.load("isNullObject");
    employeesName.load("isNullObject");


    await context.sync();


    // ----------------------------------------------------------
    // ПЕРЕВІРКА
    // ----------------------------------------------------------

    if (categoriesName.isNullObject) {
      throw new Error(
        "Не знайдено іменований діапазон RBD_Categories."
      );
    }


    if (citiesName.isNullObject) {
      throw new Error(
        "Не знайдено іменований діапазон RBD_Cities."
      );
    }


    if (employeesName.isNullObject) {
      throw new Error(
        "Не знайдено іменований діапазон RBD_Employees."
      );
    }


    // ----------------------------------------------------------
    // ОТРИМУЄМО ДІАПАЗОНИ
    // ----------------------------------------------------------

    const categoriesRange =
      categoriesName.getRange();

    const citiesRange =
      citiesName.getRange();

    const employeesRange =
      employeesName.getRange();


    categoriesRange.load("values");
    citiesRange.load("values");
    employeesRange.load("values");


    await context.sync();


    // ----------------------------------------------------------
    // ЗАПОВНЮЄМО HTML SELECT
    // ----------------------------------------------------------

    fillSelect(
      "category",
      categoriesRange.values,
      "Оберіть категорію"
    );


    fillSelect(
      "city",
      citiesRange.values,
      "Оберіть місто"
    );


    fillSelect(
      "executor",
      employeesRange.values,
      "Оберіть виконавця"
    );

  });

}


// ============================================================
// ЗАПОВНЕННЯ ВИПАДАЮЧОГО СПИСКУ
// ============================================================

function fillSelect(
  elementId,
  values,
  placeholder
) {

  const select =
    document.getElementById(
      elementId
    );


  if (!select) {
    return;
  }


  // Очищаємо старі значення
  select.innerHTML = "";


  // Перший рядок
  const firstOption =
    document.createElement(
      "option"
    );


  firstOption.value = "";

  firstOption.textContent =
    placeholder;


  select.appendChild(
    firstOption
  );


  // Унікальні значення
  const uniqueValues =
    new Set();


  values.forEach((row) => {

    if (!row || row.length === 0) {
      return;
    }


    const value =
      String(
        row[0] ?? ""
      ).trim();


    if (!value) {
      return;
    }


    uniqueValues.add(value);

  });


  uniqueValues.forEach((value) => {

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
