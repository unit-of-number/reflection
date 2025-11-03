// ====== 設定項目 ======
const MASTER_SHEET_ID = "1PCV7K_dWb2EH1hcUzh4HIN42wIobKwexABpyDDiy4MY"; // ★ご自身の教師用シートのID
const LOCAL_SHEET_NAME = "ふりかえり記録";
const CONFIG_SHEET_NAME = "_config";
// =====================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ふりかえり入力')
    .addItem('今日のふりかえりを入力・提出する', 'showUnitDialog')
    .addToUi();
}

function showUnitDialog() {
  const htmlTemplate = HtmlService.createTemplateFromFile('dialog');
  htmlTemplate.pastEntries = getPastReflections();
  htmlTemplate.lastUnitInfo = getLastUnitInfo(); 
  
  const htmlOutput = htmlTemplate.evaluate()
      .setWidth(1800)
      .setHeight(1200);
      
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '今日のふりかえりを入力しよう');
}

function getLastUnitInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    return { unitName: "", mainGoal: "" };
  }
  const data = configSheet.getRange("A1:B1").getValues();
  return {
    unitName: data[0][0],
    mainGoal: data[0][1]
  };
}

function getPastReflections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOCAL_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    entries.push({
      timestamp: data[i][0], unitName:  data[i][1], mainGoal:  data[i][2],
      date:      data[i][3], lesson:    data[i][4], text:      data[i][5],
      feedback:  data[i][6] || ""
    });
  }
  return entries.reverse();
}

function processReflection(data) {
  if (!data || !data.text || data.text.trim() === "") {
    return "エラー: ふりかえりが入力されていません。";
  }
  
  try {
    const studentSs = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- ★★★ ここが唯一の、しかし最も重要な修正点です ★★★ ---
    // ファイル名から、" - " より前の部分（＝児童名）だけを正確に抜き出す
    const studentName = studentSs.getName().split(" - ")[0];
    
    const timestamp = new Date();
    
    // --- 1. 新しい単元情報を記憶する ---
    let configSheet = studentSs.getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) {
      configSheet = studentSs.insertSheet(CONFIG_SHEET_NAME);
      configSheet.hideSheet();
    }
    configSheet.getRange("A1").setValue(data.unitName);
    configSheet.getRange("B1").setValue(data.mainGoal);

    // --- 2. 児童用の記録シートにデータを書き込む ---
    let localSheet = studentSs.getSheetByName(LOCAL_SHEET_NAME);
    if (!localSheet) {
      localSheet = studentSs.insertSheet(LOCAL_SHEET_NAME);
      localSheet.appendRow(["提出日時", "単元名", "単元を貫く課題", "日付", "時間", "ふりかえり", "先生からのコメント"]);
      localSheet.hideSheet();
    }
    const rowData = [timestamp, data.unitName, data.mainGoal, data.date, data.lesson, data.text];
    localSheet.appendRow(rowData);
    
    // --- 3. 教師用の集約シートにデータを書き込む ---
    const masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheets()[0];
    masterSheet.appendRow([timestamp, studentName, data.unitName, data.mainGoal, data.date, data.lesson, data.text]);
    
    return "提出が完了しました！";
  } catch (e) {
    Logger.log("エラー発生: " + e.message);
    return "エラーが発生しました。先生に連絡してください。";
  }
}

function recordFeedback(timestampString, reflectionText, feedbackComment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "ふりかえり記録";
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) { return; }
  
  if (sheet.getRange("G1").getValue() !== "先生からのコメント") {
    sheet.getRange("G1").setValue("先生からのコメント");
  }

  const data = sheet.getDataRange().getValues();
  const targetTimestamp = new Date(timestampString).getTime();

  for (let i = 1; i < data.length; i++) {
    const rowTimestamp = new Date(data[i][0]).getTime();
    const rowReflection = data[i][5];

    if (rowTimestamp === targetTimestamp && rowReflection === reflectionText) {
      sheet.getRange(i + 1, 7).setValue(feedbackComment);
      return;
    }
  }
}
