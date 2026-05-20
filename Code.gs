/**
 * 設定：シート名と列の定義
 */
const CONFIG = {
  SHEET_NAME: "タスク",
  DONE_SHEET_NAME: "完了タスク",
  PROJECT_SHEET_NAME: "プロジェクト",

  // 新しい列配置
  COL_ID: 1,           // A列: タスクID
  COL_STATUS: 2,       // B列: 完了
  COL_REPEAT: 3,       // C列: Repeat
  COL_TYPE: 4,         // D列: 項目 (Private/Biz)
  COL_PROJECT_NAME: 5, // E列: プロジェクト名 (ARRAYFORMULA数式)
  COL_PROJECT: 6,      // F列: プロジェクトID
  COL_TASK: 7,         // G列: タスク
  COL_DATE_E: 8,       // H列: 期日
  COL_START_TIME: 9,   // I列: 開始時間
  COL_END_TIME: 10,    // J列: 終了時間
  COL_DESCRIPTION: 11, // K列: 詳細
  COL_SORT_ORDER: 13,  // M列: 順序
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

  if (row < 2) return; // ヘッダー行の編集は無視

  // --- 0. タスクID等のデフォルト値自動割り当て ---
  if (sheetName === CONFIG.SHEET_NAME && val !== "") {
    const idRange = sheet.getRange(row, CONFIG.COL_ID);
    if (!idRange.getValue()) {
      const newId = Utilities.getUuid().split('-')[0];
      idRange.setValue(newId);

      // 完了状態(B列)が空なら「FALSE」を自動入力
      const statusRange = sheet.getRange(row, CONFIG.COL_STATUS);
      if (statusRange.getValue() === "") {
        statusRange.setValue("FALSE");
      }

      // Repeat(C列)が空なら「None」を自動入力
      const repeatRange = sheet.getRange(row, CONFIG.COL_REPEAT);
      if (repeatRange.getValue() === "") {
        repeatRange.setValue("None");
      }

      // 項目(D列)が空なら「Private」を自動入力
      const typeRange = sheet.getRange(row, CONFIG.COL_TYPE);
      if (typeRange.getValue() === "") {
        typeRange.setValue("Private");
      }

      // 順序(M列)が空なら最大順序+1を自動入力
      const sortOrderRange = sheet.getRange(row, CONFIG.COL_SORT_ORDER);
      if (sortOrderRange.getValue() === "") {
        sortOrderRange.setValue(getMaxSortOrder(sheet) + 1);
      }
    }
  }

  // --- 1. タスク完了（B列チェック）時の処理 ---
  if (sheetName === CONFIG.SHEET_NAME && col === CONFIG.COL_STATUS) {
    if (val === true || String(val).toUpperCase() === "TRUE") {
      if (e.oldValue && String(e.oldValue).toUpperCase() === "TRUE") return;
      processTaskCompletion(ss, sheet, row);
      return;
    }
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
  // A列から一番右の列まで全てのデータを取得する
  const lastCol = sheet.getLastColumn();
  const rowRange = sheet.getRange(row, 1, 1, lastCol);
  const rowData = rowRange.getValues()[0];

  const repeatType = rowData[CONFIG.COL_REPEAT - 1];

  // カレンダー登録とリピート作成を分離し、一方のエラーで止まらないように安全に処理
  // --- カレンダー＆リピート作成 ---
  try {
    syncToCalendarLogic(rowData);
  } catch (err) {
    console.error("カレンダー作成エラー: " + err.message);
  }

  try {
    const rtTypeStr = String(repeatType || "").trim();
    if (rtTypeStr !== "" && rtTypeStr.toUpperCase() !== "NONE") {
      handleRepeatTask(sheet, rowData, rtTypeStr);
    }
  } catch (err) {
    console.error("リピート作成エラー: " + err.message);
  }

  // --- 完了転記と削除 ---
  try {
    const doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME) || ss.insertSheet(CONFIG.DONE_SHEET_NAME);

    // ARRAYFORMULAの影響を受けないように本当の最終行を取得
    const lastRowDone = getRealLastRow(doneSheet) + 1;

    // 転記の配列の長さを列数に合わせる
    doneSheet.getRange(lastRowDone, 1, 1, rowData.length).setValues([rowData]);

    // 完了タスクシートのプロジェクト名の列(F列)もクリアしてあげる
    doneSheet.getRange(lastRowDone, CONFIG.COL_PROJECT_NAME).clearContent();

    doneSheet.getRange(lastRowDone, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
    doneSheet.getRange(lastRowDone, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
    doneSheet.getRange(lastRowDone, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");

    SpreadsheetApp.flush();
    sheet.deleteRow(row);

    ss.toast("タスクを完了しました。", "処理完了");

    // 完了タスクが移動・削除された後、自動で並び替えを実行する
    sortAll();

  } catch (err) {
    ss.toast("転記/削除エラー: " + err.message, "エラー");
  }
}

/**
 * リピートタスクを末尾に新規作成
 */
function handleRepeatTask(sheet, rowData, repeatType) {
  const currentDate = new Date(rowData[CONFIG.COL_DATE_E - 1]);
  if (isNaN(currentDate.getTime())) {
    SpreadsheetApp.getActiveSpreadsheet().toast("【警告】期日が日付形式ではないか、入力されていないため、次回のタスクが作られませんでした。", "エラー");
    return;
  }

  const nextDate = new Date(currentDate.getTime());
  const rType = String(repeatType).toUpperCase();
  if (rType === "DAILY") nextDate.setDate(nextDate.getDate() + 1);
  else if (rType === "WEEKLY") nextDate.setDate(nextDate.getDate() + 7);
  else if (rType === "BIWEEKLY") nextDate.setDate(nextDate.getDate() + 14);
  else if (rType === "MONTHLY") nextDate.setMonth(nextDate.getMonth() + 1);

  const newRowData = [...rowData];
  newRowData[CONFIG.COL_ID - 1] = Utilities.getUuid().split('-')[0]; // 新しいIDを発行
  newRowData[CONFIG.COL_PROJECT_NAME - 1] = ""; 
  newRowData[CONFIG.COL_DATE_E - 1] = nextDate; // 期日
  newRowData[CONFIG.COL_STATUS - 1] = "FALSE";  // 完了フラグを外す（文字列のFALSE）

  // 開始時間と終了時間のコピーは不要のためクリアする
  newRowData[CONFIG.COL_START_TIME - 1] = ""; 
  newRowData[CONFIG.COL_END_TIME - 1] = "";

  // 本当にデータが有る行の下に追加
  const targetRow = getRealLastRow(sheet) + 1;
  sheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);

  // ARRAYFORMULAの自動展開を妨げないように、追加された行のプロジェクト名(F列)のセルを空っぽに戻します
  sheet.getRange(targetRow, CONFIG.COL_PROJECT_NAME).clearContent();

  sheet.getRange(targetRow, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
  sheet.getRange(targetRow, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
  sheet.getRange(targetRow, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");
}

/**
 * ARRAYFORMULAなどで空白に見えるセルを除き、A列で本当に値が入っている最後の行を取得する関数
 */
function getRealLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 1;

  const values = sheet.getRange(1, 1, lastRow, 1).getValues(); // A列を取得
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1; // 値が入っている最後の行番号を返す
    }
  }
  return 1; // データがなければ1行目を返す
}

/**
 * 現在登録されているタスクの最大順序番号を取得する
 */
function getMaxSortOrder(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  
  const lastCol = sheet.getLastColumn();
  if (lastCol < CONFIG.COL_SORT_ORDER) return 0;
  
  const values = sheet.getRange(2, CONFIG.COL_SORT_ORDER, lastRow - 1, 1).getValues();
  let max = 0;
  for (let i = 0; i < values.length; i++) {
    const val = parseInt(values[i][0], 10);
    if (!isNaN(val) && val > max) {
      max = val;
    }
  }
  return max;
}

/**
 * カレンダー登録ロジック
 */
function syncToCalendarLogic(rowData) {
  const projectName = rowData[CONFIG.COL_PROJECT_NAME - 1];
  const taskName = rowData[CONFIG.COL_TASK - 1];
  const startDate = rowData[CONFIG.COL_DATE_E - 1]; // カレンダー基準日は「期日」
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

    // プロジェクト名が存在する列の一括置換
    const range = targetSheet.getRange(2, CONFIG.COL_PROJECT_NAME, lastRow - 1, 1);
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
  ui.createMenu('並び替え')
    .addItem('全表示', 'sortAll')
    .addToUi();
}

function sortAndFilterTasks(filterKeyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();

  const lastRow = sheet.getLastRow();
  // M列（順序）まで確実にソート範囲に含めるように拡張
  const lastCol = Math.max(sheet.getLastColumn(), CONFIG.COL_SORT_ORDER);
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  range.sort([
    // 優先順位：完了(B列) -> 期日(H列) -> 順序(M列) -> 開始時間(I列)
    {column: CONFIG.COL_STATUS, ascending: true}, 
    {column: CONFIG.COL_DATE_E, ascending: true}, 
    {column: CONFIG.COL_SORT_ORDER, ascending: true}, 
    {column: CONFIG.COL_START_TIME, ascending: true}  
  ]);

  const filterRange = sheet.getRange(1, 1, lastRow, lastCol);
  const filter = filterRange.createFilter();

  if (filterKeyword) {
    const criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(filterKeyword).build();
    filter.setColumnFilterCriteria(CONFIG.COL_TYPE, criteria);
  }
}

function sortAll() { sortAndFilterTasks(null); }

/**
 * Webアプリのエントリポイント（HTML配信）
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('タスク・プロジェクト管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * アプリ起動時の初期データ（タスク、プロジェクト、完了履歴）を一括取得
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. プロジェクトデータ取得
  const projectSheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
  const projects = [];
  if (projectSheet) {
    const values = projectSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] !== "") {
        projects.push({
          id: String(values[i][0]),
          name: String(values[i][1]),
          type: String(values[i][2])
        });
      }
    }
  }
  
  // 2. タスクデータ取得 (未完了タスク)
  const taskSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const tasks = [];
  if (taskSheet) {
    const values = taskSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][CONFIG.COL_ID - 1] !== "") {
        tasks.push(parseRowToTask(values[i]));
      }
    }
  }
  
  // 3. 完了タスクデータ取得 (直近50件)
  const doneSheet = ss.getSheetByName(CONFIG.DONE_SHEET_NAME);
  const doneTasks = [];
  if (doneSheet) {
    const values = doneSheet.getDataRange().getValues();
    const startRow = Math.max(1, values.length - 50);
    for (let i = values.length - 1; i >= startRow; i--) {
      if (values[i][CONFIG.COL_ID - 1] !== "") {
        doneTasks.push(parseRowToTask(values[i]));
      }
    }
  }
  
  return {
    tasks: tasks,
    projects: projects,
    doneTasks: doneTasks
  };
}

/**
 * 行データをフロントエンド用のタスクオブジェクトに変換
 */
function parseRowToTask(row) {
  return {
    id: String(row[CONFIG.COL_ID - 1]),
    status: String(row[CONFIG.COL_STATUS - 1]).toUpperCase() === "TRUE",
    repeat: String(row[CONFIG.COL_REPEAT - 1] || "None"),
    type: String(row[CONFIG.COL_TYPE - 1] || "Private"),
    projectName: String(row[CONFIG.COL_PROJECT_NAME - 1] || ""),
    projectId: String(row[CONFIG.COL_PROJECT - 1] || ""),
    task: String(row[CONFIG.COL_TASK - 1] || ""),
    date: formatDate(row[CONFIG.COL_DATE_E - 1]),
    startTime: formatTime(row[CONFIG.COL_START_TIME - 1]),
    endTime: formatTime(row[CONFIG.COL_END_TIME - 1]),
    description: String(row[CONFIG.COL_DESCRIPTION - 1] || "")
  };
}

/**
 * 日付オブジェクトを yyyy-MM-dd フォーマットの文字列に変換
 */
function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return "";
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(dateVal || "");
}

/**
 * 時間オブジェクトを HH:mm フォーマットの文字列に変換
 */
function formatTime(timeVal) {
  if (timeVal instanceof Date) {
    if (isNaN(timeVal.getTime())) return "";
    return Utilities.formatDate(timeVal, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(timeVal || "");
}

/**
 * アプリから新規タスクを追加
 */
function addNewTaskFromApp(taskData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error("タスクシートが見つかりません。");
  
  const newId = Utilities.getUuid().split('-')[0];
  const row = getRealLastRow(sheet) + 1;
  
  const dateVal = taskData.date ? new Date(taskData.date.replace(/-/g, "/")) : "";
  const startTimeVal = taskData.startTime ? parseTimeStringToDate(taskData.date, taskData.startTime) : "";
  const endTimeVal = taskData.endTime ? parseTimeStringToDate(taskData.date, taskData.endTime) : "";
  
  // 順序列（M列：13列目）まで考慮して配列の長さを13にする
  const newRow = Array(CONFIG.COL_SORT_ORDER).fill("");
  newRow[CONFIG.COL_ID - 1] = newId;
  newRow[CONFIG.COL_STATUS - 1] = "FALSE";
  newRow[CONFIG.COL_REPEAT - 1] = taskData.repeat || "None";
  newRow[CONFIG.COL_TYPE - 1] = taskData.type || "Private";
  newRow[CONFIG.COL_PROJECT_NAME - 1] = ""; // ARRAYFORMULA用
  newRow[CONFIG.COL_PROJECT - 1] = taskData.projectId || "";
  newRow[CONFIG.COL_TASK - 1] = taskData.task || "";
  newRow[CONFIG.COL_DATE_E - 1] = dateVal;
  newRow[CONFIG.COL_START_TIME - 1] = startTimeVal;
  newRow[CONFIG.COL_END_TIME - 1] = endTimeVal;
  newRow[CONFIG.COL_DESCRIPTION - 1] = taskData.description || "";
  newRow[CONFIG.COL_SORT_ORDER - 1] = getMaxSortOrder(sheet) + 1;
  
  sheet.getRange(row, 1, 1, newRow.length).setValues([newRow]);
  
  sheet.getRange(row, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
  sheet.getRange(row, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
  sheet.getRange(row, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");
  
  SpreadsheetApp.flush();
  sortAll();
  
  return getInitialData();
}

/**
 * 日付文字列と時間文字列からDateオブジェクトを生成
 */
function parseTimeStringToDate(dateString, timeString) {
  if (!dateString || !timeString) return "";
  const d = new Date(dateString.replace(/-/g, "/"));
  const parts = timeString.split(":");
  d.setHours(parseInt(parts[0], 10), parseInt(parts[1], 10), 0, 0);
  return d;
}

/**
 * アプリからタスクを編集
 */
function updateTaskFromApp(taskData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error("タスクシートが見つかりません。");
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("タスクがありません。");
  
  const ids = sheet.getRange(2, CONFIG.COL_ID, lastRow - 1, 1).getValues();
  let foundRow = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(taskData.id)) {
      foundRow = i + 2;
      break;
    }
  }
  
  if (foundRow === -1) throw new Error("指定されたタスクが見つかりません。ID: " + taskData.id);
  
  const dateVal = taskData.date ? new Date(taskData.date.replace(/-/g, "/")) : "";
  const startTimeVal = taskData.startTime ? parseTimeStringToDate(taskData.date, taskData.startTime) : "";
  const endTimeVal = taskData.endTime ? parseTimeStringToDate(taskData.date, taskData.endTime) : "";
  
  sheet.getRange(foundRow, CONFIG.COL_REPEAT).setValue(taskData.repeat || "None");
  sheet.getRange(foundRow, CONFIG.COL_TYPE).setValue(taskData.type || "Private");
  sheet.getRange(foundRow, CONFIG.COL_PROJECT).setValue(taskData.projectId || "");
  sheet.getRange(foundRow, CONFIG.COL_TASK).setValue(taskData.task || "");
  sheet.getRange(foundRow, CONFIG.COL_DATE_E).setValue(dateVal);
  sheet.getRange(foundRow, CONFIG.COL_START_TIME).setValue(startTimeVal);
  sheet.getRange(foundRow, CONFIG.COL_END_TIME).setValue(endTimeVal);
  sheet.getRange(foundRow, CONFIG.COL_DESCRIPTION).setValue(taskData.description || "");
  
  sheet.getRange(foundRow, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
  sheet.getRange(foundRow, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
  sheet.getRange(foundRow, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");
  
  SpreadsheetApp.flush();
  sortAll();
  
  return getInitialData();
}

/**
 * アプリからタスクを完了処理
 */
function completeTaskFromApp(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error("タスクシートが見つかりません。");
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("タスクがありません。");
  
  const ids = sheet.getRange(2, CONFIG.COL_ID, lastRow - 1, 1).getValues();
  let foundRow = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(taskId)) {
      foundRow = i + 2;
      break;
    }
  }
  
  if (foundRow === -1) throw new Error("指定されたタスクが見つかりません。ID: " + taskId);
  
  sheet.getRange(foundRow, CONFIG.COL_STATUS).setValue(true);
  processTaskCompletion(ss, sheet, foundRow);
  
  return getInitialData();
}

/**
 * アプリから新規プロジェクトを追加
 */
function addNewProjectFromApp(projectData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.PROJECT_SHEET_NAME);
  if (!sheet) throw new Error("プロジェクトシートが見つかりません。");
  
  const lastRow = sheet.getLastRow();
  let newId = 100;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => Number(row[0])).filter(val => !isNaN(val));
    if (ids.length > 0) {
      newId = Math.max(...ids) + 1;
    }
  }
  
  const targetRow = lastRow + 1;
  sheet.getRange(targetRow, 1).setValue(newId);
  sheet.getRange(targetRow, 2).setValue(projectData.name);
  sheet.getRange(targetRow, 3).setValue(projectData.type || "Private");
  
  SpreadsheetApp.flush();
  return getInitialData();
}

/**
 * アプリからタスクを削除する（カレンダー同期やリピートは行わず、行を削除）
 */
function deleteTaskFromApp(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error("タスクシートが見つかりません。");
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("タスクがありません。");
  
  const ids = sheet.getRange(2, CONFIG.COL_ID, lastRow - 1, 1).getValues();
  let foundRow = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(taskId)) {
      foundRow = i + 2;
      break;
    }
  }
  
  if (foundRow === -1) throw new Error("指定されたタスクが見つかりません。ID: " + taskId);
  
  sheet.deleteRow(foundRow);
  SpreadsheetApp.flush();
  sortAll();
  
  return getInitialData();
}

/**
 * アプリからのドラッグ＆ドロップによる並び順変更をスプレッドシートに同期する
 */
function updateSortOrderFromApp(idList) {
  if (!idList || !Array.isArray(idList)) throw new Error("無効なIDリストです。");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error("タスクシートが見つかりません。");
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return getInitialData();
  
  // IDと行番号のマップを作成
  const ids = sheet.getRange(2, CONFIG.COL_ID, lastRow - 1, 1).getValues();
  const idToRowMap = {};
  for (let i = 0; i < ids.length; i++) {
    idToRowMap[String(ids[i][0])] = i + 2;
  }
  
  // 順序（M列）を更新
  for (let index = 0; index < idList.length; index++) {
    const taskId = String(idList[index]);
    const row = idToRowMap[taskId];
    if (row) {
      sheet.getRange(row, CONFIG.COL_SORT_ORDER).setValue(index + 1);
    }
  }
  
  SpreadsheetApp.flush();
  sortAll();
  
  return getInitialData();
}

