/**
 * 設定：シート名と列の定義
 */
const CONFIG = {
  SHEET_NAME: "タスク",
  DONE_SHEET_NAME: "完了タスク",
  PROJECT_SHEET_NAME: "プロジェクト",
  COL_TYPE: 2,         // B列
  COL_PROJECT: 3,      // C列
  COL_TASK: 4,         // D列
  COL_DATE_E: 5,       // E列
  COL_DATE_F: 6,       // F列
  COL_START_TIME: 7,   // G列
  COL_END_TIME: 8,     // H列
  COL_DESCRIPTION: 9,  // I列
  COL_STATUS: 10,      // J列
  COL_REPEAT: 11,      // K列
};

/**
 * 編集を検知して実行するトリガー用関数
 */
function onEditTrigger(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const range = e.range;
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();
  const val = range.getValue();

  // --- 1. タスク完了（J列チェック）時の処理 ---
  if (sheetName === CONFIG.SHEET_NAME && col === CONFIG.COL_STATUS && val === true) {
    if (e.oldValue && String(e.oldValue).toLowerCase() === "true") return;
    processTaskCompletion(ss, sheet, row);
    return;
  }

  // --- 2. プロジェクト名変更の同期処理 ---
  if (sheetName === CONFIG.PROJECT_SHEET_NAME && col === 2) {
    const oldName = e.oldValue;
    const newName = e.value;
    if (oldName && newName && oldName !== newName) {
      syncProjectNameChange(ss, oldName, newName);
    }
  }
}

/**
 * タスク完了時のメイン処理
 */
function processTaskCompletion(ss, sheet, row) {
  const rowRange = sheet.getRange(row, 1, 1, CONFIG.COL_REPEAT);
  const rowData = rowRange.getValues()[0];
  const repeatType = rowData[CONFIG.COL_REPEAT - 1];

  // --- カレンダー＆リピート作成 ---
  try {
    syncToCalendarLogic(rowData);
    
    if (repeatType && repeatType !== "None") {
      handleRepeatTask(sheet, rowData, repeatType);
    }
  } catch (err) {
    console.error("カレンダー/リピート作成エラー: " + err.message);
  }

  // --- 完了転記と削除 ---
  try {
    const doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME) || ss.insertSheet(CONFIG.DONE_SHEET_NAME);
    doneSheet.appendRow(rowData);
    
    const lastRowDone = doneSheet.getLastRow();
    // ★修正：「MM/dd」を「M/d」に変更して、ゼロ埋めをなくしました
    doneSheet.getRange(lastRowDone, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");   // E列
    doneSheet.getRange(lastRowDone, CONFIG.COL_DATE_F).setNumberFormat("M/d(ddd)");   // F列
    doneSheet.getRange(lastRowDone, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");  // G列
    doneSheet.getRange(lastRowDone, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");    // H列

    SpreadsheetApp.flush();
    sheet.deleteRow(row);
    
    ss.toast("タスクを完了しました。", "処理完了");

  } catch (err) {
    ss.toast("転記/削除エラー: " + err.message, "エラー");
  }
}

/**
 * リピートタスクを末尾に新規作成
 */
function handleRepeatTask(sheet, rowData, repeatType) {
  const currentDate = new Date(rowData[CONFIG.COL_DATE_E - 1]);
  if (isNaN(currentDate.getTime())) return;

  const nextDate = new Date(currentDate.getTime());
  if (repeatType === "Daily") nextDate.setDate(nextDate.getDate() + 1);
  else if (repeatType === "Weekly") nextDate.setDate(nextDate.getDate() + 7);
  else if (repeatType === "Monthly") nextDate.setMonth(nextDate.getMonth() + 1);

  const newRowData = [...rowData];
  newRowData[CONFIG.COL_DATE_E - 1] = nextDate; // 期日
  newRowData[CONFIG.COL_DATE_F - 1] = nextDate; // 実施日
  newRowData[CONFIG.COL_STATUS - 1] = false;    // 完了のチェックを外す

  sheet.appendRow(newRowData);
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, CONFIG.COL_STATUS).insertCheckboxes();
  
  // ★修正：ここも「MM/dd」を「M/d」に変更しました
  sheet.getRange(lastRow, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");   // E列
  sheet.getRange(lastRow, CONFIG.COL_DATE_F).setNumberFormat("M/d(ddd)");   // F列
  sheet.getRange(lastRow, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");  // G列
  sheet.getRange(lastRow, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");    // H列
}

/**
 * カレンダー登録ロジック
 */
function syncToCalendarLogic(rowData) {
  const projectName = rowData[CONFIG.COL_PROJECT - 1];
  const taskName = rowData[CONFIG.COL_TASK - 1];
  const startDate = rowData[CONFIG.COL_DATE_F - 1];
  const startTime = rowData[CONFIG.COL_START_TIME - 1];
  const endTime = rowData[CONFIG.COL_END_TIME - 1];
  const description = rowData[CONFIG.COL_DESCRIPTION - 1];

  const title = [projectName, taskName].filter(Boolean).join(" ");
  const start = combineDateTime(startDate, startTime);
  const end = combineDateTime(startDate, endTime);

  if (start && end) {
    CalendarApp.getDefaultCalendar().createEvent(title, start, end, { description: description });
  }
}

/**
 * プロジェクト名の一括置換
 */
function syncProjectNameChange(ss, oldName, newName) {
  const sheetsToUpdate = [CONFIG.SHEET_NAME, CONFIG.DONE_SHEET_NAME];
  
  sheetsToUpdate.forEach(name => {
    const targetSheet = ss.getSheetByName(name);
    if (!targetSheet) return;

    const lastRow = targetSheet.getLastRow();
    if (lastRow < 2) return;
    
    const range = targetSheet.getRange(2, CONFIG.COL_PROJECT, lastRow - 1, 1);
    const values = range.getValues();
    const updatedValues = values.map(row => [row[0] === oldName ? newName : row[0]]);
    range.setValues(updatedValues);
  });
}

/**
 * 補助関数：日付と時間を結合
 */
function combineDateTime(date, time) {
  if (!(date instanceof Date)) return null;
  const dateTime = new Date(date.getTime());
  if (time instanceof Date) {
    dateTime.setHours(time.getHours(), time.getMinutes(), 0);
  } else {
    dateTime.setHours(0, 0, 0);
  }
  return dateTime;
}

// --- メニューとソート機能 ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('管理メニュー')
    .addItem('並び替え（Private）', 'sortPrivate')
    .addItem('並び替え（Biz）', 'sortBiz')
    .addSeparator()
    .addItem('並び替え（全表示）', 'sortAll')
    .addToUi();
}

function sortAndFilterTasks(filterKeyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;
  
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;
  
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  range.sort([
    {column: CONFIG.COL_STATUS, ascending: true}, 
    {column: CONFIG.COL_DATE_E, ascending: true}, 
    {column: CONFIG.COL_START_TIME, ascending: true}  
  ]);
  
  const filterRange = sheet.getRange(1, 1, lastRow, lastCol);
  const filter = filterRange.createFilter();
  
  if (filterKeyword) {
    const criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(filterKeyword).build();
    filter.setColumnFilterCriteria(CONFIG.COL_TYPE, criteria);
  }
}

function sortPrivate() { sortAndFilterTasks("Private"); }
function sortBiz()     { sortAndFilterTasks("Biz"); }
function sortAll()     { sortAndFilterTasks(null); }
