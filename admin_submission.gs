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
      
      // 名前を設定
      updateUserNamesInSheet(newSheet, userIdToNameMap);
    }
    
    // ヘッダーと日付を設定
    updateSheetHeader(newSheet, year, month);
    
    // 平日停止日の斜線を設定
    applyDiagonalLinesForClosedDays(newSheet, year, month);
    
  } catch (e) {
    console.error('新しいシートの初期化中にエラーが発生しました: ' + e.message);
  }
  
  console.log(`新しいシート ${newSheetName} を作成しました。`);
}

/**
 * 毎日18:00に実行される関数（トリガー関数）
 * 当月のシートに予約データを更新
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

    // シートのユーザーIDに対応する名前を設定
    updateUserNamesInSheet(targetSheet, userIdToNameMap);

    // 月次の予約データを取得
    const reservationData = getMonthlyReservationCounts(year, month);
    if (!reservationData.success) {
      console.error("予約データの取得に失敗しました:", reservationData.message);
      return;
    }

    // ヘッダー情報（タイトルと日付）を更新
    updateSheetHeader(targetSheet, year, month);

    // 平日停止日の斜線を設定
    applyDiagonalLinesForClosedDays(targetSheet, year, month);

    // 前半・後半ブロックごとにユーザーIDと行のマッピングを作成
    const userRowMap_1_16 = createUserRowMap(targetSheet, 5, 37);
    const userRowMap_17_31 = createUserRowMap(targetSheet, 45, 77);

    // 既存の予約データをクリア（3列目以降）
    clearExistingReservationData(targetSheet);

    // 予約データを書き込む
    const dataToUpdate = [];
    const { breakfast: breakfastReservations, dinner: dinnerReservations } = reservationData;

    processReservations(breakfastReservations, false, userRowMap_1_16, userRowMap_17_31, dataToUpdate);
    processReservations(dinnerReservations, true, userRowMap_1_16, userRowMap_17_31, dataToUpdate);
    
    // データを一括更新
    dataToUpdate.forEach(data => {
      targetSheet.getRange(data.row, data.col).setValue(data.value);
    });
    
    console.log(`${sheetName} の予約データを更新しました。`);

  } catch (e) {
    console.error('updateDailyMealSheet Error: ' + e.message + " Stack: " + e.stack);
  }
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
 */
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  // 毎月1日00:00のトリガー
  ScriptApp.newTrigger('createMonthlySheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
  
  // 毎日18:00のトリガー
  ScriptApp.newTrigger('updateDailyMealSheet')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
  
  console.log('トリガーを設定しました。');
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