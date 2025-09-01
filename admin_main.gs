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

// ==========================================
// フロントエンド用ラッパー関数群
// Google Apps Scriptでは、フロントエンドから直接アクセスできるのは
// admin_main.gsの関数のみなので、他のファイルの関数をここでラップする
// ==========================================

/**
 * フロントエンド用：メニューリスト取得
 */
function getMenuLists() {
  return getMenuListsImpl();
}

/**
 * フロントエンド用：メニュー更新
 */
function updateMenuForCalendar(calendarId, mealType, menuName, calorieValue, year, month) {
  return updateMenuForCalendarImpl(calendarId, mealType, menuName, calorieValue, year, month);
}

/**
 * フロントエンド用：月次予約データを取得
 */
function getMonthlyReservationCounts(year, month) {
  return getMonthlyReservationCountsImpl(year, month);
}

/**
 * フロントエンド用：食事原紙作成
 */
function createMonthlyMealSheet(year, month) {
  return createMonthlyMealSheetImpl(year, month);
}

/**
 * フロントエンド用：食事原紙URL取得
 */
function getMealSheetUrl() {
  return getMealSheetUrlImpl();
}

/**
 * フロントエンド用：募集停止データ取得
 */
function getRecruitmentStops(year, month) {
  return getRecruitmentStopsImpl(year, month);
}

/**
 * フロントエンド用：募集停止切り替え
 */
function toggleRecruitmentStop(date, mealType, year, month) {
  return toggleRecruitmentStopImpl(date, mealType, year, month);
}

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
