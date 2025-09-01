/**
 * 食事原紙を確認するためのスプレッドシートURLを取得する
 * @return {Object} スプレッドシートのURLを含むオブジェクト
 */
function getMealSheetUrl() {
  const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
  const mealSS = SpreadsheetApp.openById(mealSheetId);
  
  return {
    success: true,
    url: mealSS.getUrl()
  };
}

/**
 * 毎月1日00:00に新しい月のシートを作成する（トリガー関数）
 */
function createMonthlySheet() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const newSheetName = `食事原紙_${yyyyMM}`;
  
  const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
  const mealSS = SpreadsheetApp.openById(mealSheetId);
  
  // 既に同名のシートがある場合は作成しない
  const existingSheet = mealSS.getSheetByName(newSheetName);
  if (existingSheet) {
    console.log(`シート ${newSheetName} は既に存在します。`);
    return;
  }
  
  // テンプレートシートを取得
  const templateSheet = mealSS.getSheetByName("食事原紙");
  if (!templateSheet) {
    console.error("テンプレートシート「食事原紙」が見つかりません。");
    return;
  }
  
  // テンプレートをコピーして新しいシートを作成
  const newSheet = templateSheet.copyTo(mealSS);
  newSheet.setName(newSheetName);
  
  // 作成したシートに初期データを設定（ユーザー名のみ）
  try {
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // ユーザーシートからIDと名前の対応表を作成
    const usersSheet = ss.getSheetByName("users");
    if (usersSheet) {
      const usersData = usersSheet.getDataRange().getValues();
      const usersHeaders = usersData[0];
      const userIdIndex = usersHeaders.indexOf("user_id");
      const userNameIndex = usersHeaders.indexOf("name");
      const userIdToNameMap = {};
      for (let i = 1; i < usersData.length; i++) {
        userIdToNameMap[usersData[i][userIdIndex]] = usersData[i][userNameIndex];
      }
      
      // 名前のみ設定（テンプレートの関数はそのまま使用）
      updateUserNamesInSheet(newSheet, userIdToNameMap);
    }
    
  } catch (e) {
    console.error('新しいシートの初期化中にエラーが発生しました: ' + e.message);
  }
  
  console.log(`新しいシート ${newSheetName} を作成しました。`);
}

/**
 * 毎日12:00に実行される関数（トリガー関数）
 * 当月のシートに予約データを更新
 * ※ テンプレートの関数はそのまま使用し、ユーザー名設定と当日の予約データのみ更新
 */
function updateDailyMealSheet() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const sheetName = `食事原紙_${yyyyMM}`;
  
  const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
  const mealSS = SpreadsheetApp.openById(mealSheetId);
  
  // 対象シートを取得
  const targetSheet = mealSS.getSheetByName(sheetName);
  if (!targetSheet) {
    console.error(`シート ${sheetName} が見つかりません。`);
    return;
  }
  
  try {
    // 予約データ用スプレッドシートからデータを取得
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // ユーザーシートからIDと名前の対応表を作成
    const usersSheet = ss.getSheetByName("users");
    if (!usersSheet) {
      console.error("「users」シートが見つかりません。");
      return;
    }
    
    const usersData = usersSheet.getDataRange().getValues();
    const usersHeaders = usersData[0];
    const userIdIndex = usersHeaders.indexOf("user_id");
    const userNameIndex = usersHeaders.indexOf("name");
    const userIdToNameMap = {};
    for (let i = 1; i < usersData.length; i++) {
      userIdToNameMap[usersData[i][userIdIndex]] = usersData[i][userNameIndex];
    }

    // ユーザー名を設定（毎回更新）
    updateUserNamesInSheet(targetSheet, userIdToNameMap);

    // 月次の予約データを取得
    const reservationData = getMonthlyReservationCounts(year, month);
    if (!reservationData.success) {
      console.error("予約データの取得に失敗しました:", reservationData.message);
      return;
    }

    // 前半・後半ブロックごとにユーザーIDと行のマッピングを作成
    const userRowMap_1_16 = createUserRowMap(targetSheet, 5, 37);
    const userRowMap_17_31 = createUserRowMap(targetSheet, 45, 77);

    // 今日の日付を取得
    const today = new Date();
    const todayStr = formatDate(today);
    const todayDayOfMonth = today.getDate();

    // 今日分のデータのみ更新
    const dataToUpdate = [];
    const { breakfast: breakfastReservations, dinner: dinnerReservations } = reservationData;

    // 朝食の今日分のデータを処理
    const todayBreakfast = breakfastReservations.find(item => item.date === todayStr);
    if (todayBreakfast && todayBreakfast.users.length > 0) {
      processSingleDayReservations(todayBreakfast, false, todayDayOfMonth, userRowMap_1_16, userRowMap_17_31, dataToUpdate);
    }

    // 夕食の今日分のデータを処理
    const todayDinner = dinnerReservations.find(item => item.date === todayStr);
    if (todayDinner && todayDinner.users.length > 0) {
      processSingleDayReservations(todayDinner, true, todayDayOfMonth, userRowMap_1_16, userRowMap_17_31, dataToUpdate);
    }
    
    // 今日分のデータを更新
    dataToUpdate.forEach(data => {
      targetSheet.getRange(data.row, data.col).setValue(data.value);
    });
    
    console.log(`${sheetName} の今日（${todayStr}）の予約データを更新しました。`);

  } catch (e) {
    console.error('updateDailyMealSheet Error: ' + e.message + " Stack: " + e.stack);
  }
}

/**
 * 単一日の予約データを処理
 */
function processSingleDayReservations(dayData, isDinner, dayOfMonth, userRowMap_1_16, userRowMap_17_31, dataToUpdate) {
  let userRowMap;
  let relativeDay;

  if (dayOfMonth <= 16) {
    // 前半ブロックの場合
    userRowMap = userRowMap_1_16;
    relativeDay = dayOfMonth;
  } else {
    // 後半ブロックの場合
    userRowMap = userRowMap_17_31;
    relativeDay = dayOfMonth - 16;
  }

  // 1から始まる相対的な日付で列を計算する
  const column = (relativeDay - 1) * 2 + (isDinner ? 4 : 3);

  dayData.users.forEach(user => {
    const userRow = userRowMap[user.userId];
    if (userRow) {
      dataToUpdate.push({row: userRow, col: column, value: 1});
    }
  });
}

/**
 * シートのユーザーIDに対応する名前を設定
 */
function updateUserNamesInSheet(sheet, userIdToNameMap) {
  // 前半ブロック (5行目〜37行目)
  updateNamesInBlock(sheet, 5, 37, userIdToNameMap);
  // 後半ブロック (45行目〜77行目)
  updateNamesInBlock(sheet, 45, 77, userIdToNameMap);
}

/**
 * 指定範囲のユーザー名を更新
 */
function updateNamesInBlock(sheet, startRow, endRow, userIdToNameMap) {
  const idRange = sheet.getRange(`A${startRow}:A${endRow}`);
  const nameRange = sheet.getRange(`B${startRow}:B${endRow}`);
  const idValues = idRange.getValues();
  const namesToSet = [];

  for (let i = 0; i < idValues.length; i++) {
    const userId = idValues[i][0];
    if (userId && !isNaN(userId)) {
      namesToSet.push([userIdToNameMap[userId] || '']);
    } else {
      namesToSet.push(['']);
    }
  }
  nameRange.setValues(namesToSet);
}

/**
 * シートのヘッダー（タイトルと日付）を更新
 */
function updateSheetHeader(sheet, year, month) {
  // タイトルを更新
  const titleRange = sheet.getRange("A1");
  const currentTitle = titleRange.getValue();
  if (currentTitle && currentTitle.toString().includes('月度')) {
    titleRange.setValue(currentTitle.toString().replace(/\d+月度/, `${month}月度`));
  }

  // 日付ヘッダーを正確に設定
  updateDateHeaders(sheet, year, month);
}

/**
 * 日付ヘッダーを正確に設定
 */
function updateDateHeaders(sheet, year, month) {
  const daysInMonth = new Date(year, month, 0).getDate();
  
  // 前半ブロック (1日〜16日)
  for (let day = 1; day <= Math.min(16, daysInMonth); day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
    
    const dayCol = (day - 1) * 2 + 3; // C列から開始
    sheet.getRange(2, dayCol).setValue(day);
    sheet.getRange(2, dayCol + 1).setValue(dayOfWeek);
  }
  
  // 後半ブロック (17日〜月末)
  for (let day = 17; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
    
    const relativeDay = day - 16;
    const dayCol = (relativeDay - 1) * 2 + 3; // 後半ブロックのC列から開始
    sheet.getRange(42, dayCol).setValue(day);
    sheet.getRange(42, dayCol + 1).setValue(dayOfWeek);
  }
}

/**
 * 平日停止日に斜線を適用
 */
function applyDiagonalLinesForClosedDays(sheet, year, month) {
  const daysInMonth = new Date(year, month, 0).getDate();
  
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = date.getDay(); // 0=日曜, 6=土曜
    
    // 土曜日は朝食・夕食ともに停止、日曜日は夕食のみ停止
    if (dayOfWeek === 6 || dayOfWeek === 0) {
      applyDiagonalLineForDay(sheet, day, dayOfWeek === 6); // 土曜日は朝食も停止
    }
  }
}

/**
 * 指定日に斜線を適用
 */
function applyDiagonalLineForDay(sheet, day, includeBreakfast) {
  let blockStartRow, relativeDay;
  
  if (day <= 16) {
    // 前半ブロック
    blockStartRow = 5;
    relativeDay = day;
  } else {
    // 後半ブロック
    blockStartRow = 45;
    relativeDay = day - 16;
  }
  
  const breakfastCol = (relativeDay - 1) * 2 + 3;
  const dinnerCol = breakfastCol + 1;
  
  // 各ユーザー行に斜線を適用
  for (let row = blockStartRow; row < blockStartRow + 33; row++) {
    if (includeBreakfast) {
      // 朝食セルに斜線
      applyDiagonalLineToCell(sheet, row, breakfastCol);
    }
    // 夕食セルに斜線
    applyDiagonalLineToCell(sheet, row, dinnerCol);
  }
}

/**
 * セルに斜線を適用
 */
function applyDiagonalLineToCell(sheet, row, col) {
  const cell = sheet.getRange(row, col);
  cell.setBorder(null, null, null, null, true, null); // 斜線を設定
  cell.setBackground('#f0f0f0'); // 薄いグレー背景
}

/**
 * ユーザーIDと行のマッピングを作成
 */
function createUserRowMap(sheet, startRow, endRow) {
  const userRowMap = {};
  const idRange = sheet.getRange(`A${startRow}:A${endRow}`);
  const idValues = idRange.getValues();

  for (let i = 0; i < idValues.length; i++) {
    const userId = idValues[i][0];
    if (userId && !isNaN(userId)) {
      userRowMap[userId] = startRow + i;
    }
  }
  return userRowMap;
}

/**
 * 既存の予約データをクリア
 */
function clearExistingReservationData(sheet) {
  const maxCol = sheet.getMaxColumns();
  if (maxCol > 2) {
    // 前半ブロック
    sheet.getRange(5, 3, 33, maxCol - 2).clearContent();
    // 後半ブロック
    sheet.getRange(45, 3, 33, maxCol - 2).clearContent();
  }
}

/**
 * 予約データを処理
 */
function processReservations(reservations, isDinner, userRowMap_1_16, userRowMap_17_31, dataToUpdate) {
  reservations.forEach(dayData => {
    if (dayData.users.length > 0) {
      const dayOfMonth = parseInt(dayData.date.split('-')[2], 10);
      
      let userRowMap;
      let relativeDay;

      if (dayOfMonth <= 16) {
        // 前半ブロックの場合
        userRowMap = userRowMap_1_16;
        relativeDay = dayOfMonth;
      } else {
        // 後半ブロックの場合
        userRowMap = userRowMap_17_31;
        relativeDay = dayOfMonth - 16;
      }

      // 1から始まる相対的な日付で列を計算する
      const column = (relativeDay - 1) * 2 + (isDinner ? 4 : 3);

      dayData.users.forEach(user => {
        const userRow = userRowMap[user.userId];
        if (userRow) {
          dataToUpdate.push({row: userRow, col: column, value: 1});
        }
      });
    }
  });
}

/**
 * トリガーを設定する関数（手動で1回実行する）
 * - 月次トリガー：毎月1日 00:00に新しい月のシート作成
 * - 日次トリガー：毎日 12:00にデータ更新
 */
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  // 毎月1日00:00のトリガー（新しい月のシート作成）
  ScriptApp.newTrigger('createMonthlySheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
  
  // 毎日12:00のトリガー（データ更新）
  ScriptApp.newTrigger('updateDailyMealSheet')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
  
  console.log('トリガーを設定しました：月次(1日00:00), 日次(毎日12:00)');
}

// ==========================================
// テスト用関数群
// ==========================================

/**
 * 【テスト用】現在の月のシートを手動作成
 * Google Apps Scriptエディタで直接実行してテストできます
 */
function testCreateCurrentMonthSheet() {
  console.log('=== 現在の月のシート作成テスト開始 ===');
  
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  console.log(`対象: ${year}年${month}月`);
  
  try {
    createMonthlySheet();
    console.log('✅ シート作成テスト完了');
  } catch (e) {
    console.error('❌ シート作成テストでエラー:', e.message);
  }
}

/**
 * 【テスト用】現在の月のシートデータを手動更新
 * Google Apps Scriptエディタで直接実行してテストできます
 */
function testUpdateCurrentMonthSheet() {
  console.log('=== 現在の月のシート更新テスト開始 ===');
  
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  console.log(`対象: ${year}年${month}月`);
  
  try {
    updateDailyMealSheet();
    console.log('✅ シート更新テスト完了');
  } catch (e) {
    console.error('❌ シート更新テストでエラー:', e.message);
  }
}

/**
 * 【テスト用】指定した年月のシートを作成
 * @param {number} testYear テスト用の年
 * @param {number} testMonth テスト用の月
 */
function testCreateSpecificMonthSheet(testYear, testMonth) {
  console.log(`=== ${testYear}年${testMonth}月のシート作成テスト開始 ===`);
  
  const yyyyMM = `${testYear}${testMonth.toString().padStart(2, "0")}`;
  const newSheetName = `食事原紙_${yyyyMM}`;
  
  const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
  const mealSS = SpreadsheetApp.openById(mealSheetId);
  
  // 既に同名のシートがある場合は削除
  const existingSheet = mealSS.getSheetByName(newSheetName);
  if (existingSheet) {
    console.log(`既存のシート ${newSheetName} を削除します。`);
    mealSS.deleteSheet(existingSheet);
  }
  
  // テンプレートシートを取得
  const templateSheet = mealSS.getSheetByName("食事原紙");
  if (!templateSheet) {
    console.error("テンプレートシート「食事原紙」が見つかりません。");
    return;
  }
  
  // テンプレートをコピーして新しいシートを作成
  const newSheet = templateSheet.copyTo(mealSS);
  newSheet.setName(newSheetName);
  
  // 作成したシートに初期データを設定
  try {
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // ユーザーシートからIDと名前の対応表を作成
    const usersSheet = ss.getSheetByName("users");
    if (usersSheet) {
      const usersData = usersSheet.getDataRange().getValues();
      const usersHeaders = usersData[0];
      const userIdIndex = usersHeaders.indexOf("user_id");
      const userNameIndex = usersHeaders.indexOf("name");
      const userIdToNameMap = {};
      for (let i = 1; i < usersData.length; i++) {
        userIdToNameMap[usersData[i][userIdIndex]] = usersData[i][userNameIndex];
      }
      
      console.log(`ユーザー数: ${Object.keys(userIdToNameMap).length}人`);
      
      // 名前を設定
      updateUserNamesInSheet(newSheet, userIdToNameMap);
      console.log('✅ ユーザー名設定完了');
    }
    
    // ヘッダーと日付を設定
    updateSheetHeader(newSheet, testYear, testMonth);
    console.log('✅ ヘッダー・日付設定完了');
    
    // 平日停止日の斜線を設定
    applyDiagonalLinesForClosedDays(newSheet, testYear, testMonth);
    console.log('✅ 斜線設定完了');
    
  } catch (e) {
    console.error('新しいシートの初期化中にエラーが発生しました: ' + e.message);
  }
  
  console.log(`✅ テストシート ${newSheetName} を作成しました。`);
  console.log(`URL: ${mealSS.getUrl()}`);
}

/**
 * 【テスト用】指定した年月のシートデータを更新
 * @param {number} testYear テスト用の年
 * @param {number} testMonth テスト用の月
 */
function testUpdateSpecificMonthSheet(testYear, testMonth) {
  console.log(`=== ${testYear}年${testMonth}月のシート更新テスト開始 ===`);
  
  const yyyyMM = `${testYear}${testMonth.toString().padStart(2, "0")}`;
  const sheetName = `食事原紙_${yyyyMM}`;
  
  const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
  const mealSS = SpreadsheetApp.openById(mealSheetId);
  
  // 対象シートを取得
  const targetSheet = mealSS.getSheetByName(sheetName);
  if (!targetSheet) {
    console.error(`シート ${sheetName} が見つかりません。先にシートを作成してください。`);
    return;
  }
  
  try {
    // 予約データ用スプレッドシートからデータを取得
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // ユーザーシートからIDと名前の対応表を作成
    const usersSheet = ss.getSheetByName("users");
    if (!usersSheet) {
      console.error("「users」シートが見つかりません。");
      return;
    }
    
    const usersData = usersSheet.getDataRange().getValues();
    const usersHeaders = usersData[0];
    const userIdIndex = usersHeaders.indexOf("user_id");
    const userNameIndex = usersHeaders.indexOf("name");
    const userIdToNameMap = {};
    for (let i = 1; i < usersData.length; i++) {
      userIdToNameMap[usersData[i][userIdIndex]] = usersData[i][userNameIndex];
    }

    console.log(`ユーザー数: ${Object.keys(userIdToNameMap).length}人`);

    // シートのユーザーIDに対応する名前を設定
    updateUserNamesInSheet(targetSheet, userIdToNameMap);
    console.log('✅ ユーザー名更新完了');

    // 月次の予約データを取得
    const reservationData = getMonthlyReservationCounts(testYear, testMonth);
    if (!reservationData.success) {
      console.error("予約データの取得に失敗しました:", reservationData.message);
      return;
    }

    console.log(`朝食予約データ: ${reservationData.breakfast.length}日分`);
    console.log(`夕食予約データ: ${reservationData.dinner.length}日分`);

    // ヘッダー情報（タイトルと日付）を更新
    updateSheetHeader(targetSheet, testYear, testMonth);
    console.log('✅ ヘッダー更新完了');

    // 平日停止日の斜線を設定
    applyDiagonalLinesForClosedDays(targetSheet, testYear, testMonth);
    console.log('✅ 斜線設定完了');

    // 前半・後半ブロックごとにユーザーIDと行のマッピングを作成
    const userRowMap_1_16 = createUserRowMap(targetSheet, 5, 37);
    const userRowMap_17_31 = createUserRowMap(targetSheet, 45, 77);

    console.log(`前半ブロックユーザー数: ${Object.keys(userRowMap_1_16).length}人`);
    console.log(`後半ブロックユーザー数: ${Object.keys(userRowMap_17_31).length}人`);

    // 既存の予約データをクリア（3列目以降）
    clearExistingReservationData(targetSheet);
    console.log('✅ 既存データクリア完了');

    // 予約データを書き込む
    const dataToUpdate = [];
    const { breakfast: breakfastReservations, dinner: dinnerReservations } = reservationData;

    processReservations(breakfastReservations, false, userRowMap_1_16, userRowMap_17_31, dataToUpdate);
    processReservations(dinnerReservations, true, userRowMap_1_16, userRowMap_17_31, dataToUpdate);
    
    console.log(`更新対象セル数: ${dataToUpdate.length}個`);
    
    // データを一括更新
    dataToUpdate.forEach(data => {
      targetSheet.getRange(data.row, data.col).setValue(data.value);
    });
    
    console.log(`✅ ${sheetName} の予約データ更新テスト完了`);
    console.log(`URL: ${mealSS.getUrl()}`);

  } catch (e) {
    console.error('updateDailyMealSheet Error: ' + e.message + " Stack: " + e.stack);
  }
}

/**
 * 【テスト用】8月のシートを作成・更新する簡単テスト
 * 一つの関数で作成から更新まで実行
 */
function testCreateAndUpdate2025August() {
  console.log('=== 2025年8月のシート作成・更新テスト ===');
  
  // 1. シート作成
  testCreateSpecificMonthSheet(2025, 8);
  
  // 2. データ更新
  Utilities.sleep(2000); // 2秒待機
  testUpdateSpecificMonthSheet(2025, 8);
  
  console.log('=== 2025年8月のテスト完了 ===');
}

/**
 * 【テスト用】9月のシートを作成・更新する簡単テスト
 * 一つの関数で作成から更新まで実行
 */
function testCreateAndUpdate2025September() {
  console.log('=== 2025年9月のシート作成・更新テスト ===');
  
  // 1. シート作成
  testCreateSpecificMonthSheet(2025, 9);
  
  // 2. データ更新
  Utilities.sleep(2000); // 2秒待機
  testUpdateSpecificMonthSheet(2025, 9);
  
  console.log('=== 2025年9月のテスト完了 ===');
}

// ==========================================
// 予約データ取得機能
// ==========================================

/**
 * 月次予約データ取得（内部実装）
 */
function getMonthlyReservationCounts(year, month) {
  try {
    console.log('=== getMonthlyReservationCountsImpl開始 ===');
    console.log('パラメータ:', year, month);
    
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    const bCalendarSheetName = "b_calendar_" + yyyyMM;
    const dCalendarSheetName = "d_calendar_" + yyyyMM;
    const bReservationSheetName = "b_reservations_" + yyyyMM;
    const dReservationSheetName = "d_reservations_" + yyyyMM;
    
    // シートの存在確認
    const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
    const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
    const bReservationSheet = ss.getSheetByName(bReservationSheetName);
    const dReservationSheet = ss.getSheetByName(dReservationSheetName);
    const usersSheet = ss.getSheetByName("users");
    const bMenuSheet = ss.getSheetByName("b_menus");
    const dMenuSheet = ss.getSheetByName("d_menus");
    
    if (!bCalendarSheet || !dCalendarSheet || !bReservationSheet || !dReservationSheet || !usersSheet) {
      return {
        success: false,
        message: '必要なシートが見つかりません。'
      };
    }
  
    // データの取得
    const bCalendarData = bCalendarSheet.getDataRange().getValues();
    const dCalendarData = dCalendarSheet.getDataRange().getValues();
    const bReservationData = bReservationSheet.getDataRange().getValues();
    const dReservationData = dReservationSheet.getDataRange().getValues();
    const usersData = usersSheet.getDataRange().getValues();
    
    // メニューデータの取得
    let bMenuData = [];
    let dMenuData = [];
    
    if (bMenuSheet) {
      bMenuData = bMenuSheet.getDataRange().getValues();
    }
    if (dMenuSheet) {
      dMenuData = dMenuSheet.getDataRange().getValues();
    }
    
    // メニューマップの作成
    const bMenuMap = {};
    if (bMenuData.length > 1) {
      const bMenuIdIndex = bMenuData[0].indexOf("b_menu_id");
      const bMenuNameIndex = bMenuData[0].indexOf("breakfast_menu");
      const bCalorieIndex = bMenuData[0].indexOf("calorie");
      
      if (bMenuIdIndex !== -1 && bMenuNameIndex !== -1) {
        for (let i = 1; i < bMenuData.length; i++) {
          const menuId = bMenuData[i][bMenuIdIndex];
          const menuName = bMenuData[i][bMenuNameIndex];
          const calorie = bCalorieIndex !== -1 ? bMenuData[i][bCalorieIndex] : 0;
          bMenuMap[menuId] = {
            name: menuName,
            calorie: calorie || 0
          };
        }
      }
    }
    
    const dMenuMap = {};
    if (dMenuData.length > 1) {
      const dMenuIdIndex = dMenuData[0].indexOf("d_menu_id");
      const dMenuNameIndex = dMenuData[0].indexOf("dinner_menu");
      const dCalorieIndex = dMenuData[0].indexOf("calorie");
      
      if (dMenuIdIndex !== -1 && dMenuNameIndex !== -1) {
        for (let i = 1; i < dMenuData.length; i++) {
          const menuId = dMenuData[i][dMenuIdIndex];
          const menuName = dMenuData[i][dMenuNameIndex];
          const calorie = dCalorieIndex !== -1 ? dMenuData[i][dCalorieIndex] : 0;
          dMenuMap[menuId] = {
            name: menuName,
            calorie: calorie || 0
          };
        }
      }
    }
    
    // ユーザーマップの作成
    const userMap = {};
    const userIdIndex = usersData[0].indexOf("user_id");
    const userNameIndex = usersData[0].indexOf("name");
    for (let i = 1; i < usersData.length; i++) {
      const userId = usersData[i][userIdIndex];
      const userName = usersData[i][userNameIndex];
      userMap[userId] = userName;
    }
    
    // カレンダーデータの処理
    const bCalendarDateMap = {};
    const bCalendarHeaders = bCalendarData[0];
    const bCalendarIdIndex = bCalendarHeaders.indexOf("b_calendar_id");
    const bCalendarDateIndex = bCalendarHeaders.indexOf("date");
    const bCalendarMenuIdIndex = bCalendarHeaders.indexOf("b_menu_id");
    
    for (let i = 1; i < bCalendarData.length; i++) {
      const calendarId = bCalendarData[i][bCalendarIdIndex];
      const date = bCalendarData[i][bCalendarDateIndex];
      const menuId = bCalendarData[i][bCalendarMenuIdIndex];
      
      if (date instanceof Date) {
        const dateStr = date.getFullYear() + '-' + 
                       (date.getMonth() + 1).toString().padStart(2, '0') + '-' + 
                       date.getDate().toString().padStart(2, '0');
        bCalendarDateMap[calendarId] = {
          date: dateStr,
          menuId: menuId,
          menuName: bMenuMap[menuId] ? bMenuMap[menuId].name : "未設定",
          calorie: bMenuMap[menuId] ? bMenuMap[menuId].calorie : 0
        };
      }
    }
    
    const dCalendarDateMap = {};
    const dCalendarHeaders = dCalendarData[0];
    const dCalendarIdIndex = dCalendarHeaders.indexOf("d_calendar_id");
    const dCalendarDateIndex = dCalendarHeaders.indexOf("date");
    const dCalendarMenuIdIndex = dCalendarHeaders.indexOf("d_menu_id");
    
    for (let i = 1; i < dCalendarData.length; i++) {
      const calendarId = dCalendarData[i][dCalendarIdIndex];
      const date = dCalendarData[i][dCalendarDateIndex];
      const menuId = dCalendarData[i][dCalendarMenuIdIndex];
      
      if (date instanceof Date) {
        const dateStr = date.getFullYear() + '-' + 
                       (date.getMonth() + 1).toString().padStart(2, '0') + '-' + 
                       date.getDate().toString().padStart(2, '0');
        dCalendarDateMap[calendarId] = {
          date: dateStr,
          menuId: menuId,
          menuName: dMenuMap[menuId] ? dMenuMap[menuId].name : "未設定",
          calorie: dMenuMap[menuId] ? dMenuMap[menuId].calorie : 0
        };
      }
    }
    
    // 予約データの処理
    const bReservationCounts = {};
    const bReservationUsers = {};
    const bReservationHeaders = bReservationData[0];
    const bReservationCalendarIdIndex = bReservationHeaders.indexOf("b_calendar_id");
    const bReservationUserIdIndex = bReservationHeaders.indexOf("user_id");
    const bReservationStatusIndex = bReservationHeaders.indexOf("is_reserved");
    
    for (let i = 1; i < bReservationData.length; i++) {
      const row = bReservationData[i];
      const calendarId = row[bReservationCalendarIdIndex];
      const userId = row[bReservationUserIdIndex];
      const isReserved = row[bReservationStatusIndex];
      
      if (isReserved) {
        if (!bReservationCounts[calendarId]) {
          bReservationCounts[calendarId] = 0;
          bReservationUsers[calendarId] = [];
        }
        
        bReservationCounts[calendarId]++;
        bReservationUsers[calendarId].push({
          userId: userId,
          userName: userMap[userId] || "Unknown"
        });
      }
    }
    
    const dReservationCounts = {};
    const dReservationUsers = {};
    const dReservationHeaders = dReservationData[0];
    const dReservationCalendarIdIndex = dReservationHeaders.indexOf("d_calendar_id");
    const dReservationUserIdIndex = dReservationHeaders.indexOf("user_id");
    const dReservationStatusIndex = dReservationHeaders.indexOf("is_reserved");
    
    for (let i = 1; i < dReservationData.length; i++) {
      const row = dReservationData[i];
      const calendarId = row[dReservationCalendarIdIndex];
      const userId = row[dReservationUserIdIndex];
      const isReserved = row[dReservationStatusIndex];
      
      if (isReserved) {
        if (!dReservationCounts[calendarId]) {
          dReservationCounts[calendarId] = 0;
          dReservationUsers[calendarId] = [];
        }
        
        dReservationCounts[calendarId]++;
        dReservationUsers[calendarId].push({
          userId: userId,
          userName: userMap[userId] || "Unknown"
        });
      }
    }
    
    // 結果の形成
    const breakfastReservations = [];
    const dinnerReservations = [];
    
    // 朝食の集計
    for (const calendarId in bCalendarDateMap) {
      const dateInfo = bCalendarDateMap[calendarId];
      breakfastReservations.push({
        calendarId: calendarId,
        date: dateInfo.date,
        menuId: dateInfo.menuId,
        menuName: dateInfo.menuName,
        calorie: dateInfo.calorie,
        count: bReservationCounts[calendarId] || 0,
        users: bReservationUsers[calendarId] || []
      });
    }
    
    // 夕食の集計
    for (const calendarId in dCalendarDateMap) {
      const dateInfo = dCalendarDateMap[calendarId];
      dinnerReservations.push({
        calendarId: calendarId,
        date: dateInfo.date,
        menuId: dateInfo.menuId,
        menuName: dateInfo.menuName,
        calorie: dateInfo.calorie,
        count: dReservationCounts[calendarId] || 0,
        users: dReservationUsers[calendarId] || []
      });
    }
    
    // 日付でソート
    breakfastReservations.sort((a, b) => a.date.localeCompare(b.date));
    dinnerReservations.sort((a, b) => a.date.localeCompare(b.date));
    
    console.log('✅ データ処理完了:', {
      breakfastCount: breakfastReservations.length,
      dinnerCount: dinnerReservations.length
    });
    
    return {
      success: true,
      year: year,
      month: month,
      breakfast: breakfastReservations,
      dinner: dinnerReservations
    };
    
  } catch (error) {
    console.error('❌ getMonthlyReservationCountsImpl エラー:', error);
    return {
      success: false,
      message: '処理中にエラーが発生しました: ' + error.message,
      breakfast: [],
      dinner: []
    };
  }
}

/**
 * 募集停止データ取得（内部実装）
 */
function getRecruitmentStops(year, month) {
  try {
    console.log('=== getRecruitmentStopsImpl開始 ===');
    
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    const recruitmentStopsSheet = ss.getSheetByName("recruitment_stops");
    if (!recruitmentStopsSheet) {
      console.log('recruitment_stopsシートが見つかりません。空のデータを返します。');
      return {
        success: true,
        stops: {}
      };
    }
    
    const data = recruitmentStopsSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        success: true,
        stops: {}
      };
    }
    
    const headers = data[0];
    const dateIndex = headers.indexOf("date");
    const breakfastIndex = headers.indexOf("is_breakfast_stopped");
    const dinnerIndex = headers.indexOf("is_dinner_stopped");
    const isActiveIndex = headers.indexOf("is_active");
    
    const stops = {};
    const targetYearMonth = year + '-' + (month < 10 ? '0' + month : month);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[dateIndex];
      const isBreakfastStopped = row[breakfastIndex];
      const isDinnerStopped = row[dinnerIndex];
      const isActive = row[isActiveIndex];
      
      if (!isActive) continue;
      
      let dateStr;
      if (date instanceof Date) {
        dateStr = date.getFullYear() + '-' + 
                 (date.getMonth() + 1).toString().padStart(2, '0') + '-' + 
                 date.getDate().toString().padStart(2, '0');
      } else {
        dateStr = date.toString();
      }
      
      if (dateStr.startsWith(targetYearMonth)) {
        stops[dateStr] = {
          breakfast: !!isBreakfastStopped,
          dinner: !!isDinnerStopped
        };
      }
    }
    
    return {
      success: true,
      stops: stops
    };
    
  } catch (error) {
    console.error('❌ getRecruitmentStopsImpl エラー:', error);
    return {
      success: false,
      message: '募集停止データの取得に失敗しました: ' + error.message,
      stops: {}
    };
  }
}

/**
 * 募集停止切り替え（内部実装）
 */
function toggleRecruitmentStop(date, mealType, year, month) {
  try {
    console.log('=== toggleRecruitmentStopImpl開始 ===');
    
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    let recruitmentStopsSheet = ss.getSheetByName("recruitment_stops");
    if (!recruitmentStopsSheet) {
      recruitmentStopsSheet = ss.insertSheet("recruitment_stops");
      recruitmentStopsSheet.appendRow([
        "date", "is_breakfast_stopped", "is_dinner_stopped", "is_active", "created_at", "updated_at"
      ]);
    }
    
    const data = recruitmentStopsSheet.getDataRange().getValues();
    const headers = data[0];
    const dateIndex = headers.indexOf("date");
    const breakfastIndex = headers.indexOf("is_breakfast_stopped");
    const dinnerIndex = headers.indexOf("is_dinner_stopped");
    const isActiveIndex = headers.indexOf("is_active");
    const updatedAtIndex = headers.indexOf("updated_at");
    
    // 既存レコードを検索
    let existingRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][dateIndex];
      let rowDateStr;
      
      if (rowDate instanceof Date) {
        rowDateStr = rowDate.getFullYear() + '-' + 
                    (rowDate.getMonth() + 1).toString().padStart(2, '0') + '-' + 
                    rowDate.getDate().toString().padStart(2, '0');
      } else {
        rowDateStr = rowDate.toString();
      }
      
      if (rowDateStr === date && data[i][isActiveIndex]) {
        existingRowIndex = i;
        break;
      }
    }
    
    const now = new Date();
    let newBreakfastStopped = false;
    let newDinnerStopped = false;
    let message = '';
    
    if (existingRowIndex !== -1) {
      const currentBreakfastStopped = data[existingRowIndex][breakfastIndex];
      const currentDinnerStopped = data[existingRowIndex][dinnerIndex];
      
      if (mealType === 'breakfast') {
        newBreakfastStopped = !currentBreakfastStopped;
        newDinnerStopped = currentDinnerStopped;
        message = newBreakfastStopped ? '朝食の募集を停止しました' : '朝食の募集停止を解除しました';
      } else {
        newBreakfastStopped = currentBreakfastStopped;
        newDinnerStopped = !currentDinnerStopped;
        message = newDinnerStopped ? '夕食の募集を停止しました' : '夕食の募集停止を解除しました';
      }
      
      recruitmentStopsSheet.getRange(existingRowIndex + 1, breakfastIndex + 1).setValue(newBreakfastStopped);
      recruitmentStopsSheet.getRange(existingRowIndex + 1, dinnerIndex + 1).setValue(newDinnerStopped);
      recruitmentStopsSheet.getRange(existingRowIndex + 1, updatedAtIndex + 1).setValue(now);
      
    } else {
      if (mealType === 'breakfast') {
        newBreakfastStopped = true;
        newDinnerStopped = false;
        message = '朝食の募集を停止しました';
      } else {
        newBreakfastStopped = false;
        newDinnerStopped = true;
        message = '夕食の募集を停止しました';
      }
      
      recruitmentStopsSheet.appendRow([
        date,
        newBreakfastStopped,
        newDinnerStopped,
        true,
        now,
        now
      ]);
    }
    
    return {
      success: true,
      message: message
    };
    
  } catch (error) {
    console.error('❌ toggleRecruitmentStopImpl エラー:', error);
    return {
      success: false,
      message: '募集停止の切り替え中にエラーが発生しました: ' + error.message
    };
  }
}

