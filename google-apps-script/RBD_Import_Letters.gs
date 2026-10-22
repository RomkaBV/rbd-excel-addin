function importRbdLetters() {
  const root = DriveApp.getFolderById("18hQb_ym1aXPD2sDIlvEBOIlntcfXhEw4");
  const it = root.getFoldersByName("Імпорт листування");
  const inbox = it.hasNext() ? it.next() : root.createFolder("Імпорт листування");
  const files = inbox.getFiles();
  let count = 0;
  while (files.hasNext()) {
    const f = files.next();
    const m = /^([A-Za-z0-9_-]+)__\d+_\d+__(.+\.(?:msg|eml))$/i.exec(f.getName());
    if (!m || f.getSize() > 4194304) continue;
    const req = root.getFoldersByName(m[1]);
    const folder = req.hasNext() ? req.next() : root.createFolder(m[1]);
    const dirs = folder.getFoldersByName("Листування");
    f.moveTo(dirs.hasNext() ? dirs.next() : folder.createFolder("Листування"));
    count++;
  }
  Logger.log("Імпортовано листів: " + count);
}
