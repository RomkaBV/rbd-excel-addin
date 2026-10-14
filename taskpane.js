/* РБД 2.0 — Excel JavaScript Add-in, no backend. */
'use strict';
const $ = id => document.getElementById(id);
const state = { people: [], statuses: [], cities: [], locations: [], categories: [], customers: [], priorities: [], requests: [], sheets: [] };
const ACTIONS = ['Створити','Призначити','Взяти в роботу','Змінити статус','Коментар','Закрити','Додати місто','Додати локацію','Додати категорію','Додати замовника'];
const C = {
  users: 'tblRBDUsers2026', requests: 'tblRBDRequests2026',
  statuses: 'tblRBDStatuses2026', cities: 'tblRBDCities2026',
  locations: 'tblRBDLocations2026', categories: 'tblRBDCategories2026',
  customers: 'tblRBDCustomers2026', priorities: 'tblRBDPriorities2026'
};
function notify(message, error = false) {
  const n = $('notice'); n.textContent = message; n.className = error ? 'err' : 'ok'; n.style.display = 'block';
}
function escapeText(v) { return String(v == null ? '' : v); }
function selectOptions(id, values, placeholder = '— Оберіть —') {
  const el = $(id); el.replaceChildren();
  const empty = document.createElement('option'); empty.value = ''; empty.textContent = placeholder; el.append(empty);
  for (const value of [...new Set(values.filter(Boolean))]) {
    const op = document.createElement('option'); op.value = value; op.textContent = value; el.append(op);
  }
}
function toObjects(header, body) {
  return body.map(row => Object.fromEntries(header.map((h, i) => [String(h), row[i] ?? ''])));
}
async function loadData() {
  $('reload').disabled = true;
  try {
    const result = await Excel.run(async context => {
      const names = Object.entries(C);
      const items = names.map(([key, name]) => ({ key, table: context.workbook.tables.getItem(name) }));
      for (const item of items) {
        item.range = item.table.getRange();
        item.range.load('values');
      }
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      const out = {};
      for (const item of items) out[item.key] = toObjects(item.range.values[0], item.range.values.slice(1));
      out.sheetNames = sheets.items.map(s => s.name);
      return out;
    });
    state.people = result.users.filter(r => r['Активний'] === 'Так');
    state.statuses = result.statuses.filter(r => r['Активний'] === 'Так');
    state.cities = result.cities.filter(r => r['Активний'] === 'Так');
    state.locations = result.locations.filter(r => r['Активний'] === 'Так');
    state.categories = result.categories.filter(r => r['Активний'] === 'Так');
    state.customers = result.customers.filter(r => r['Активний'] === 'Так');
    state.priorities = result.priorities.filter(r => r['Активний'] === 'Так');
    state.requests = result.requests;
    state.sheets = result.sheetNames;
    const prev = $('person').value;
    selectOptions('person', state.people.map(r => r['ПІБ']), '— Кабінет —');
    if (state.people.some(r => r['ПІБ'] === prev)) $('person').value = prev;
    else if (state.people.length) $('person').value = state.people[0]['ПІБ'];
    selectOptions('action', ACTIONS, '— Виберіть дію —');
    selectOptions('executor', state.people.map(r => r['ПІБ']), '— Без виконавця —');
    selectOptions('status', state.statuses.map(r => r['Назва статусу']));
    selectOptions('city', state.cities.map(r => r['Назва']));
    selectOptions('category', state.categories.map(r => r['Назва']));
    selectOptions('customer', state.customers.map(r => r['Назва']));
    selectOptions('priority', state.priorities.map(r => r['Назва']));
    $('priority').value = 'Середній';
    updateLocations(); render();
    notify('Дані книги завантажено.');
  } catch (err) { notify('Не вдалося прочитати книгу: ' + err.message + '\nПеревір назви таблиць на вкладці «Налаштування надбудови».', true); }
  finally { $('reload').disabled = false; }
}
function updateLocations() {
  const city = $('city').value;
  const old = $('location').value;
  const locations = state.locations.filter(l => !city || l['Місто'] === city).map(l => l['Назва']);
  selectOptions('location', locations);
  if (locations.includes(old)) $('location').value = old;
}
function render() {
  const person = $('person').value;
  const requests = state.requests.filter(r => r['ID заявки']);
  const finished = ['Закрита', 'Скасована', 'Виконана'];
  $('total').textContent = requests.length;
  $('active').textContent = requests.filter(r => !finished.includes(r['Статус'])).length;
  $('finished').textContent = requests.filter(r => finished.includes(r['Статус'])).length;
  $('mine').textContent = requests.filter(r => r['Виконавець'] === person).length;
  const body = $('requestRows'); body.replaceChildren();
  for (const row of requests.slice(-30).reverse()) {
    const tr = document.createElement('tr');
    for (const key of ['ID заявки','Тема заявки','Статус','Виконавець']) {
      const td = document.createElement('td'); td.textContent = escapeText(row[key]); tr.appendChild(td);
    }
    tr.addEventListener('click', () => $('requestId').value = escapeText(row['ID заявки']));
    body.appendChild(tr);
  }
  if (!requests.length) { const tr = document.createElement('tr'); const td = document.createElement('td'); td.colSpan = 4; td.textContent = 'Заявок поки немає'; tr.appendChild(td); body.appendChild(tr); }
}
function cabinetTab() {
  const name = $('person').value;
  const person = state.people.find(r => r['ПІБ'] === name);
  if (!person) throw new Error('Спочатку обери виконавця.');
  const surname = name.trim().split(/\s+/)[0];
  // Manager has a personal cabinet too; the add-in deliberately uses personal cabinet.
  if (state.sheets.includes(surname)) return surname;
  const older = ['Кабінет_' + person['ID працівника'], 'РБД_' + person['ID працівника']];
  const found = older.find(tab => state.sheets.includes(tab));
  if (found) return found;
  throw new Error('Не знайдено персональну вкладку виконавця: ' + surname);
}
async function openCabinet() {
  try {
    const name = cabinetTab();
    await Excel.run(async context => {
      context.workbook.worksheets.getItem(name).activate();
      await context.sync();
    });
    notify('Відкрито вкладку «' + name + '».');
  } catch (err) { notify(err.message, true); }
}
async function saveCommand() {
  $('save').disabled = true;
  try {
    const tab = cabinetTab();
    const action = $('action').value;
    if (!action) throw new Error('Оберіть дію.');
    if (action === 'Створити' && !$('subject').value.trim()) throw new Error('Вкажіть тему заявки.');
    if (!['Створити','Додати місто','Додати локацію','Додати категорію','Додати замовника'].includes(action) && !$('requestId').value.trim()) throw new Error('Для цієї дії необхідний ID заявки.');
    const person = state.people.find(r => r['ПІБ'] === $('person').value);
    if (action === 'Призначити' && person && !person['Роль'].includes('Керівник')) {
      const selected = $('executor').value;
      if (selected && selected !== person['ПІБ']) throw new Error('Призначати іншого може лише керівник.');
    }
    const cid = 'ADD-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).slice(2,8).toUpperCase();
    const fields = {
      'ID команди': cid, 'Дія': action, 'ID заявки': $('requestId').value.trim(),
      'Тема заявки': $('subject').value.trim(), 'Опис заявки': $('description').value.trim(),
      'Виконавець': $('executor').value, 'Статус': $('status').value,
      'Місто': $('city').value, 'Локація': $('location').value,
      'Категорія': $('category').value, 'Замовник': $('customer').value,
      'Пріоритет': $('priority').value, 'Коментар': $('comment').value.trim(),
      'Очікувана версія': '', 'Стан обробки': 'Нова', 'Результат': ''
    };
    await Excel.run(async context => {
      const tables = context.workbook.worksheets.getItem(tab).tables;
      tables.load('items/name');
      await context.sync();
      let input = null;
      for (const t of tables.items) {
        const head = t.getHeaderRowRange(); head.load('values');
      }
      await context.sync();
      for (const t of tables.items) {
        const hs = t.getHeaderRowRange().values[0];
        if (hs.includes('ID команди') && hs.includes('Стан обробки')) { input = t; break; }
      }
      if (!input) throw new Error('На вкладці «' + tab + '» немає таблиці команд.');
      const cols = input.getHeaderRowRange().values[0];
      input.rows.add(null, [cols.map(c => fields[c] ?? '')], true);
      await context.sync();
    });
    notify('Команду ' + cid + ' записано у вкладку «' + tab + '».\nТепер запусти основний Office Script РБД 2.0 для її обробки.');
  } catch (err) { notify('Не вдалося записати команду: ' + err.message, true); }
  finally { $('save').disabled = false; }
}
Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) { notify('Ця надбудова призначена для Excel.', true); return; }
  $('reload').addEventListener('click', loadData);
  $('openSheet').addEventListener('click', openCabinet);
  $('person').addEventListener('change', render);
  $('city').addEventListener('change', updateLocations);
  $('save').addEventListener('click', saveCommand);
  loadData();
});
