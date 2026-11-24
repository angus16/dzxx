// 獨自修仙：接收遊戲上傳的資料。從試算表的擴充功能建立時寫進那份試算表；
// 若是獨立的 Apps Script 專案，會自動建立試算表「獨自修仙資料」（先執行一次 setup 可看到網址）
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = getSS();
    upsert(ss, '局紀錄', d.gameCols, d.gameRows, true);
    upsert(ss, '回合紀錄', d.roundCols, d.roundRows, false);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

function clean(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

// 同一個 game_id：局紀錄會覆蓋成最新一筆（補上問卷），回合紀錄只寫一次
function upsert(ss, name, cols, rows, overwrite) {
  if (!cols || !rows || !rows.length) return;
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(cols);
  const last = sh.getLastRow();
  const id = rows[0][0];
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]) : [];
  const at = ids.indexOf(id);
  const data = rows.map(r => cols.map((c, i) => clean(r[i])));
  if (at >= 0) {
    if (overwrite) sh.getRange(at + 2, 1, 1, cols.length).setValues([data[0]]);
    return;
  }
  sh.getRange(last + 1, 1, data.length, cols.length).setValues(data);
}

function doGet() {
  return ContentService.createTextOutput('獨自修仙資料收集中');
}

function getSS() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('獨自修仙資料');
  p.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// 手動執行一次：完成授權，並在記錄裡印出資料寫入的試算表網址
function setup() {
  Logger.log(getSS().getUrl());
}
