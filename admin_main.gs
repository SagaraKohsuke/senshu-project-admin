function doGet() {
  return HtmlService.createTemplateFromFile("admin_index")
    .evaluate()
    .setTitle("泉州会館 管理人画面");
}

const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
const ss = SpreadsheetApp.openById(spreadsheetId);

// 新しい食事原紙用スプレッドシート
const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
const mealSS = SpreadsheetApp.openById(mealSheetId);

/**
 * フロントエンド接続テスト用の関数
 */
function testConnection() {
  return {
    success: true,
    message: "フロントエンドとの接続は正常です",
    timestamp: new Date().toString(),
    gasVersion: "Google Apps Script",
    currentTime: Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss")
  };
}
