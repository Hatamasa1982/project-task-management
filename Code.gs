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
  COL_PROJECT: 5,      // E列: プロジェクトID
  COL_PROJECT_NAME: 6, // F列: プロジェクト名 (ARRAYFORMULA)
  COL_TASK: 7,         // G列: タスク
  COL_DATE_E: 8,       // H列: 期日
  COL_DATE_F: 9,       // I列: 実施日
  COL_START_TIME: 10,  // J列: 開始時間
  COL_END_TIME: 11,    // K列: 終了時間
  COL_DESCRIPTION: 12, // L列: 詳細
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
    doneSheet.getRange(lastRowDone, CONFIG.COL_DATE_F).setNumberFormat("M/d(ddd)");
    doneSheet.getRange(lastRowDone, CONFIG.COL_START_TIME).setNumberFormat("hh:mm");
    doneSheet.getRange(lastRowDone, CONFIG.COL_END_TIME).setNumberFormat("hh:mm");

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
  newRowData[CONFIG.COL_DATE_F - 1] = nextDate; // 実施日
  newRowData[CONFIG.COL_STATUS - 1] = "FALSE";  // 完了フラグを外す（文字列のFALSE）

  // 本当にデータが有る行の下に追加
  const targetRow = getRealLastRow(sheet) + 1;
  sheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
  
  // ARRAYFORMULAの自動展開を妨げないように、追加された行のプロジェクト名(F列)のセルを空っぽに戻します
  sheet.getRange(targetRow, CONFIG.COL_PROJECT_NAME).clearContent();
  
  sheet.getRange(targetRow, CONFIG.COL_DATE_E).setNumberFormat("M/d(ddd)");
  sheet.getRange(targetRow, CONFIG.COL_DATE_F).setNumberFormat("M/d(ddd)");
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
 * カレンダー登録ロジック
 */
function syncToCalendarLogic(rowData) {
  const projectName = rowData[CONFIG.COL_PROJECT_NAME - 1];
  const taskName = rowData[CONFIG.COL_TASK - 1];
  const startDate = rowData[CONFIG.COL_DATE_F - 1]; // カレンダー基準日は「実施日」
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
    // 今回の構成に伴い、優先順位は B列(完了フラグ) -> H列(期日) -> J列(開始時間) となる
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
