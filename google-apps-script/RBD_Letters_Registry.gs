/**
 * Створює приватний реєстр листів, що вже знаходяться в Google Drive/RBD.
 * Виконується вручну у тому самому Apps Script проєкті.
 * Жодних публічних посилань чи дозволів не створює.
 */
function rebuildRbdLettersRegistry() {
  const root = DriveApp.getFolderById("18hQb_ym1aXPD2sDIlvEBOIlntcfXhEw4");
  const props = PropertiesService.getScriptProperties();
  let ss;
  const existingId = props.getProperty("RBD_LETTERS_REGISTRY_ID");
  if (existingId) {
    ss = SpreadsheetApp.openById(existingId);
  } else {
    ss = SpreadsheetApp.create("RBD_Letters_Registry");
    const file = DriveApp.getFileById(ss.getId());
    file.moveTo(root);
    props.setProperty("RBD_LETTERS_REGISTRY_ID", ss.getId());
  }
  const sheet = ss.getSheetByName("Letters") || ss.insertSheet("Letters");
  const records = [["ID заявки","ID файла","Назва файла","Тип","Розмір, байт","Дата створення","Посилання Google Drive"]];
  const requests = root.getFolders();
  while (requests.hasNext()) {
    const request = requests.next();
    if (!/^RBD-[A-Za-z0-9_-]+$/.test(request.getName())) continue;
    const correspondence = request.getFoldersByName("Листування");
    while (correspondence.hasNext()) {
      const folder = correspondence.next();
      const files = folder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        if (!/\.(eml|msg)$/i.test(file.getName())) continue;
        records.push([
          request.getName(),file.getId(),file.getName(),
          /\.msg$/i.test(file.getName())?"MSG":"EML",
          file.getSize(),file.getDateCreated(),file.getUrl()
        ]);
      }
    }
  }
  sheet.clearContents();
  sheet.getRange(1,1,records.length,records[0].length).setValues(records);
  sheet.setFrozenRows(1);
  Logger.log("Реєстр листів: "+ss.getUrl());
  Logger.log("Знайдено листів: "+(records.length-1));
  return {url:ss.getUrl(),count:records.length-1};
}
