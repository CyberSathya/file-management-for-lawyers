// Weekly backup of the rack spreadsheet.
// Lives in its own Apps Script project (not the File Management web app), so staff using the
// app are never asked for Google Drive access.
//
// Setup: script.google.com -> New project -> paste this file -> put your spreadsheet's ID
// below (the long part of its URL between /d/ and /edit) -> Save -> run setupBackup once.

const SPREADSHEET_ID = 'PASTE_SPREADSHEET_ID_HERE';
const BACKUP_FOLDER = 'Rack Backups';
const BACKUP_NAME = 'Racks backup';
const KEEP = 12;   // weekly copies kept (about 3 months); older ones go to Drive's Bin

// Makes a dated copy, e.g. "Racks backup 2026-11-08", then trims old copies.
function backupNow() {
  if (/PASTE_/.test(SPREADSHEET_ID)) throw new Error('Put your spreadsheet ID into SPREADSHEET_ID first.');
  const folder = backupFolder_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const copy = DriveApp.getFileById(SPREADSHEET_ID).makeCopy(BACKUP_NAME + ' ' + stamp, folder);

  const copies = [];
  const it = folder.getFiles();
  while (it.hasNext()) copies.push(it.next());
  copies.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  copies.slice(KEEP).forEach(function (f) { f.setTrashed(true); });   // recoverable from Bin for 30 days

  Logger.log('Backed up to "' + BACKUP_FOLDER + '/' + copy.getName() + '". Keeping ' + Math.min(copies.length, KEEP) + ' copies.');
}

// One-time: schedules backupNow every Sunday 2-3 am (the script's time zone) and makes the first copy now.
function setupBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupNow').timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
  backupNow();
}

function backupFolder_() {
  const found = DriveApp.getFoldersByName(BACKUP_FOLDER);
  return found.hasNext() ? found.next() : DriveApp.createFolder(BACKUP_FOLDER);
}
