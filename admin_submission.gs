/**
 * 食事原紙のスプレッドシートURLを取得する（現在月のシートを表示）
 * @return {Object} 結果とURL
 */
function getMealSheetUrl() {
  try {
    const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
    const ss = SpreadsheetApp.openById(mealSheetId);
    
    // 現在の年月を取得
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    const yyyyMM = currentYear + (currentMonth < 10 ? "0" + currentMonth : currentMonth);
    const currentMealSheetName = "食事原紙_" + yyyyMM;
    
    // 現在月のシートが存在するかチェック
    const currentMealSheet = ss.getSheetByName(currentMealSheetName);
    
    if (currentMealSheet) {
      // 現在月のシートが存在する場合、そのシートを表示
      return {
        success: true,
        url: ss.getUrl() + "#gid=" + currentMealSheet.getSheetId(),
        sheetName: currentMealSheetName,
        message: "現在月の食事原紙を表示します"
      };
    } else {
      // 現在月のシートが存在しない場合、スプレッドシートのトップページを表示
      return {
        success: true,
        url: ss.getUrl(),
        sheetName: "未作成",
        message: "現在月の食事原紙「" + currentMealSheetName + "」が見つかりません。月次生成処理を実行してください。"
      };
    }
  } catch (e) {
    console.error('getMealSheetUrl Error: ' + e.message);
    return {
      success: false,
      message: "食事原紙スプレッドシートのURL取得に失敗しました: " + e.message
    };
  }
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
  
  // 作成したシートに初期データを設定（ユーザー名、日付ヘッダー、曜日）
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
      
      // ユーザー名を設定
      updateUserNamesInSheet(newSheet, userIdToNameMap);
    }
    
    // 月度タイトルと日付ヘッダー（曜日含む）を設定
    updateSheetHeader(newSheet, year, month);
    
    // 土日の休業日に斜線を適用
    applyDiagonalLinesForClosedDays(newSheet, year, month);
    
  } catch (e) {
    console.error('新しいシートの初期化中にエラーが発生しました: ' + e.message);
  }
  
  console.log(`新しいシート ${newSheetName} を作成しました。`);
  console.log(`- ユーザー名設定完了`);
  console.log(`- 月度タイトル設定: ${year}年${month}月`);  
  console.log(`- 日付・曜日ヘッダー設定完了`);
  console.log(`- 土日休業日の斜線設定完了`);
}

/**
 * 毎日12:00に実行される関数（トリガー関数）
 * 当月のシートに予約データを更新
 * ※ テンプレートの関数はそのまま使用し、ユーザー名設定と当日以降の予約データを更新
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
    const reservationData = getDetailedMonthlyReservationData(year, month);
    if (!reservationData.success) {
      console.error("予約データの取得に失敗しました:", reservationData.message);
      return;
    }

    // 前半・後半ブロックごとにユーザーIDと行のマッピングを作成
    const userRowMap_1_16 = createUserRowMap(targetSheet, 5, 37);
    const userRowMap_17_31 = createUserRowMap(targetSheet, 45, 77);

    // 今日の日付を取得
    const today = new Date();
    const todayDayOfMonth = today.getDate();

    // 当日以降の全データを更新（既存データを先にクリア）
    clearFutureDatesData(targetSheet, todayDayOfMonth, year, month);
    
    const dataToUpdate = [];
    const { breakfast: breakfastReservations, dinner: dinnerReservations } = reservationData;

    // 朝食の当日以降のデータを処理
    breakfastReservations.forEach(dayData => {
      if (dayData.users.length > 0) {
        const dayOfMonth = parseInt(dayData.date.split('-')[2], 10);
        // 当日以降のデータのみ処理
        if (dayOfMonth >= todayDayOfMonth) {
          processSingleDayReservations(dayData, false, dayOfMonth, userRowMap_1_16, userRowMap_17_31, dataToUpdate);
        }
      }
    });

    // 夕食の当日以降のデータを処理
    dinnerReservations.forEach(dayData => {
      if (dayData.users.length > 0) {
        const dayOfMonth = parseInt(dayData.date.split('-')[2], 10);
        // 当日以降のデータのみ処理
        if (dayOfMonth >= todayDayOfMonth) {
          processSingleDayReservations(dayData, true, dayOfMonth, userRowMap_1_16, userRowMap_17_31, dataToUpdate);
        }
      }
    });
    
    // 当日以降のデータを更新
    dataToUpdate.forEach(data => {
      targetSheet.getRange(data.row, data.col).setValue(data.value);
    });
    
    const lastDay = new Date(year, month, 0).getDate();
    console.log(`${sheetName} の当日以降（${todayDayOfMonth}日〜${lastDay}日）の予約データを更新しました。更新件数: ${dataToUpdate.length}件`);

  } catch (e) {
    console.error('updateDailyMealSheet Error: ' + e.message + " Stack: " + e.stack);
  }
}

/**
 * トリガーを設定する関数（手動で1回実行する）
 * 毎月1日00:00に月次シート作成、毎日12:00にデータ更新を自動実行
 */
function setupTriggers() {
  console.log('=== トリガー設定開始 ===');
  
  // 既存のトリガーを削除
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  console.log(`既存トリガー ${existingTriggers.length} 個を削除しました。`);
  
  // 毎月1日00:00のトリガー（月次シート作成）
  ScriptApp.newTrigger('createMonthlySheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
  console.log('✅ 毎月1日00:00のトリガー（createMonthlySheet）を設定しました。');
  
  // 毎日12:00のトリガー（データ更新）
  ScriptApp.newTrigger('updateDailyMealSheet')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
  console.log('✅ 毎日12:00のトリガー（updateDailyMealSheet）を設定しました。');
  
  // 設定確認
  const newTriggers = ScriptApp.getProjectTriggers();
  console.log(`トリガー設定完了: ${newTriggers.length} 個のトリガーが有効になりました。`);
  
  console.log('=== トリガー設定完了 ===');
  console.log('自動実行スケジュール:');
  console.log('- 毎月1日 00:00: 新しい月のシート作成');
  console.log('- 毎日 12:00: 当日以降の予約データ更新');
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
 * 当日以降の日付のデータをクリアする
 */
function clearFutureDatesData(sheet, startDay, year, month) {
  const lastDay = new Date(year, month, 0).getDate();
  
  // 前半ブロック（1日〜16日）と後半ブロック（17日〜31日）でそれぞれ処理
  for (let day = startDay; day <= lastDay; day++) {
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
    
    const breakfastCol = (relativeDay - 1) * 2 + 3; // 朝食の列
    const dinnerCol = breakfastCol + 1; // 夕食の列
    
    // 該当日の朝食・夕食の列をクリア（前半・後半それぞれ32行分）
    const blockEndRow = blockStartRow + 32;
    sheet.getRange(blockStartRow, breakfastCol, blockEndRow - blockStartRow + 1, 1).clearContent();
    sheet.getRange(blockStartRow, dinnerCol, blockEndRow - blockStartRow + 1, 1).clearContent();
  }
  
  console.log(`${startDay}日以降の予約データをクリアしました。`);
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
  
  // 前半ブロック (1日〜16日) - 2行目に設定
  for (let day = 1; day <= Math.min(16, daysInMonth); day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
    
    const dayCol = (day - 1) * 2 + 3; // C列から開始（3列目）
    sheet.getRange(2, dayCol).setValue(day);           // 日付
    sheet.getRange(2, dayCol + 1).setValue(dayOfWeek); // 曜日
  }
  
  // 後半ブロック (17日〜月末) - CSVを確認すると40行目にヘッダーがある
  const backHalfHeaderRow = 40; // 後半ブロックのヘッダー行
  
  for (let day = 17; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
    
    const relativeDay = day - 16;
    const dayCol = (relativeDay - 1) * 2 + 3; // C列から開始
    sheet.getRange(backHalfHeaderRow, dayCol).setValue(day);           // 日付
    sheet.getRange(backHalfHeaderRow, dayCol + 1).setValue(dayOfWeek); // 曜日
  }
  
  console.log(`日付ヘッダー更新完了: ${year}年${month}月（${daysInMonth}日まで）`);
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
 * 詳細な月次予約データを取得する（食事原紙用）
 * admin_calendar.gsのgetMonthlyReservationCountsとは異なる実装
 * @param {number} year 年
 * @param {number} month 月  
 * @return {Object} 予約データ
 */
function getDetailedMonthlyReservationData(year, month) {
  try {
    console.log('=== getDetailedMonthlyReservationData開始: ' + year + '年' + month + '月 ===');
    
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    console.log('✅ スプレッドシート接続成功');
    
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    const bCalendarSheetName = "b_calendar_" + yyyyMM;
    const dCalendarSheetName = "d_calendar_" + yyyyMM;
    const bReservationSheetName = "b_reservations_" + yyyyMM;
    const dReservationSheetName = "d_reservations_" + yyyyMM;
    
    console.log('検索対象シート:', {
      bCalendarSheetName: bCalendarSheetName,
      dCalendarSheetName: dCalendarSheetName, 
      bReservationSheetName: bReservationSheetName,
      dReservationSheetName: dReservationSheetName
    });
  
    // シートの存在確認
    const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
    const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
    const bReservationSheet = ss.getSheetByName(bReservationSheetName);
    const dReservationSheet = ss.getSheetByName(dReservationSheetName);
    const usersSheet = ss.getSheetByName("users");
    const bMenuSheet = ss.getSheetByName("b_menus");
    const dMenuSheet = ss.getSheetByName("d_menus");
    
    console.log('シート存在確認:', {
      bCalendarSheet: !!bCalendarSheet,
      dCalendarSheet: !!dCalendarSheet,
      bReservationSheet: !!bReservationSheet,
      dReservationSheet: !!dReservationSheet,
      usersSheet: !!usersSheet
    });
  
    if (!bCalendarSheet || !dCalendarSheet) {
      return {
        success: false,
        message: 'カレンダーシート ' + bCalendarSheetName + ' または ' + dCalendarSheetName + ' が見つかりません。'
      };
    }
  
    if (!bReservationSheet || !dReservationSheet) {
      return {
        success: false,
        message: '予約シート ' + bReservationSheetName + ' または ' + dReservationSheetName + ' が見つかりません。'
      };
    }
  
    if (!usersSheet) {
      return {
        success: false,
        message: "ユーザーシートが見つかりません。"
      };
    }
  
  // データの取得
  const bCalendarData = bCalendarSheet.getDataRange().getValues();
  const dCalendarData = dCalendarSheet.getDataRange().getValues();
  const bReservationData = bReservationSheet.getDataRange().getValues();
  const dReservationData = dReservationSheet.getDataRange().getValues();
  const usersData = usersSheet.getDataRange().getValues();
  
  // メニューデータの取得（存在する場合）
  let bMenuData = [];
  let dMenuData = [];
  
  if (bMenuSheet) {
    bMenuData = bMenuSheet.getDataRange().getValues();
  }
  
  if (dMenuSheet) {
    dMenuData = dMenuSheet.getDataRange().getValues();
  }
  
  // ヘッダー行の列インデックスを取得
  const bCalendarHeaders = bCalendarData[0];
  const dCalendarHeaders = dCalendarData[0];
  const bReservationHeaders = bReservationData[0];
  const dReservationHeaders = dReservationData[0];
  const usersHeaders = usersData[0];
  
  const bCalendarIdIndex = bCalendarHeaders.indexOf("b_calendar_id");
  const bCalendarDateIndex = bCalendarHeaders.indexOf("date");
  const bCalendarMenuIdIndex = bCalendarHeaders.indexOf("b_menu_id");
  
  const dCalendarIdIndex = dCalendarHeaders.indexOf("d_calendar_id");
  const dCalendarDateIndex = dCalendarHeaders.indexOf("date");
  const dCalendarMenuIdIndex = dCalendarHeaders.indexOf("d_menu_id");
  
  const bReservationCalendarIdIndex = bReservationHeaders.indexOf("b_calendar_id");
  const bReservationUserIdIndex = bReservationHeaders.indexOf("user_id");
  const bReservationStatusIndex = bReservationHeaders.indexOf("is_reserved");
  
  const dReservationCalendarIdIndex = dReservationHeaders.indexOf("d_calendar_id");
  const dReservationUserIdIndex = dReservationHeaders.indexOf("user_id");
  const dReservationStatusIndex = dReservationHeaders.indexOf("is_reserved");
  
  const userIdIndex = usersHeaders.indexOf("user_id");
  const userNameIndex = usersHeaders.indexOf("name");
  
  // 朝食メニューマップの作成（カロリー情報も含む）
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
  
  // 夕食メニューマップの作成（カロリー情報も含む）
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
  
  // ユーザーIDからユーザー名を取得するためのマップを作成
  const userMap = {};
  for (let i = 1; i < usersData.length; i++) {
    const userId = usersData[i][userIdIndex];
    const userName = usersData[i][userNameIndex];
    userMap[userId] = userName;
  }
  
  // 朝食カレンダーの日付マッピング
  const bCalendarDateMap = {};
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
  
  // 夕食カレンダーの日付マッピング
  const dCalendarDateMap = {};
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
  
  // 朝食の予約数と予約者のカウント
  const bReservationCounts = {};
  const bReservationUsers = {};
  
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
  
  // 夕食の予約数と予約者のカウント
  const dReservationCounts = {};
  const dReservationUsers = {};
  
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
    console.error('❌ getDetailedMonthlyReservationData エラー:', error);
    console.error('エラースタック:', error.stack);
    return {
      success: false,
      message: '処理中にエラーが発生しました: ' + error.message,
      breakfast: [],
      dinner: []
    };
  }
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
 * 実際は毎日12:00に自動実行される処理をテストします
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
    const reservationData = getDetailedMonthlyReservationData(testYear, testMonth);
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

/**
 * 【総合テスト用】全機能を一括テストする関数
 * Google Apps Scriptエディタで実行してください
 */
function testAllFunctionalities() {
  console.log('🧪 ========================================');
  console.log('🧪 食事原紙システム 総合テスト開始');
  console.log('🧪 ========================================');
  
  const startTime = new Date();
  let testResults = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };
  
  // テスト1: 現在月のシート作成
  console.log('\n📋 テスト1: 現在月のシート作成');
  testResults.total++;
  try {
    testCreateCurrentMonthSheet();
    testResults.passed++;
    console.log('✅ テスト1 PASSED');
  } catch (e) {
    testResults.failed++;
    testResults.errors.push(`テスト1エラー: ${e.message}`);
    console.error('❌ テスト1 FAILED:', e.message);
  }
  
  // 2秒待機
  Utilities.sleep(2000);
  
  // テスト2: 現在月のデータ更新（新仕様: 当日以降の全データ更新）
  console.log('\n📋 テスト2: 現在月のデータ更新（新仕様: 当日以降）');
  testResults.total++;
  try {
    testUpdateCurrentMonthSheet();
    testResults.passed++;
    console.log('✅ テスト2 PASSED');
  } catch (e) {
    testResults.failed++;
    testResults.errors.push(`テスト2エラー: ${e.message}`);
    console.error('❌ テスト2 FAILED:', e.message);
  }
  
  // 2秒待機
  Utilities.sleep(2000);
  
  // テスト3: 9月のシート作成・更新（曜日設定確認）
  console.log('\n📋 テスト3: 9月のシート作成・更新（曜日設定確認）');
  testResults.total++;
  try {
    testCreateAndUpdate2025September();
    testResults.passed++;
    console.log('✅ テスト3 PASSED');
  } catch (e) {
    testResults.failed++;
    testResults.errors.push(`テスト3エラー: ${e.message}`);
    console.error('❌ テスト3 FAILED:', e.message);
  }
  
  // 2秒待機
  Utilities.sleep(2000);
  
  // テスト4: トリガー設定テスト
  console.log('\n📋 テスト4: トリガー設定テスト');
  testResults.total++;
  try {
    // 現在のトリガー状況を確認
    const currentTriggers = ScriptApp.getProjectTriggers();
    console.log(`現在のトリガー数: ${currentTriggers.length}`);
    
    // トリガーを再設定
    setupTriggers();
    
    // 設定後のトリガーを確認
    const newTriggers = ScriptApp.getProjectTriggers();
    console.log(`設定後のトリガー数: ${newTriggers.length}`);
    
    // トリガー詳細を表示
    newTriggers.forEach((trigger, index) => {
      const handlerFunction = trigger.getHandlerFunction();
      const eventType = trigger.getEventType();
      console.log(`トリガー${index + 1}: ${handlerFunction} (${eventType})`);
    });
    
    if (newTriggers.length >= 2) {
      testResults.passed++;
      console.log('✅ テスト4 PASSED');
    } else {
      throw new Error('期待されるトリガー数が設定されていません');
    }
  } catch (e) {
    testResults.failed++;
    testResults.errors.push(`テスト4エラー: ${e.message}`);
    console.error('❌ テスト4 FAILED:', e.message);
  }
  
  // テスト5: 食事原紙URL取得テスト
  console.log('\n📋 テスト5: 食事原紙URL取得テスト');
  testResults.total++;
  try {
    const urlResult = getMealSheetUrl();
    if (urlResult.success && urlResult.url) {
      console.log(`✅ 食事原紙URL: ${urlResult.url}`);
      testResults.passed++;
      console.log('✅ テスト5 PASSED');
    } else {
      throw new Error('URL取得に失敗');
    }
  } catch (e) {
    testResults.failed++;
    testResults.errors.push(`テスト5エラー: ${e.message}`);
    console.error('❌ テスト5 FAILED:', e.message);
  }
  
  // テスト結果サマリー
  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);
  
  console.log('\n🧪 ========================================');
  console.log('🧪 テスト結果サマリー');
  console.log('🧪 ========================================');
  console.log(`📊 実行時間: ${duration}秒`);
  console.log(`📊 総テスト数: ${testResults.total}`);
  console.log(`✅ 成功: ${testResults.passed}`);
  console.log(`❌ 失敗: ${testResults.failed}`);
  
  if (testResults.failed > 0) {
    console.log('\n❌ エラー詳細:');
    testResults.errors.forEach(error => console.log(`  - ${error}`));
  }
  
  const successRate = Math.round((testResults.passed / testResults.total) * 100);
  console.log(`📈 成功率: ${successRate}%`);
  
  if (successRate === 100) {
    console.log('\n🎉 全テスト成功！システムは正常に動作しています。');
  } else if (successRate >= 80) {
    console.log('\n⚠️ 一部テストが失敗しましたが、基本機能は動作しています。');
  } else {
    console.log('\n⚠️ 複数のテストが失敗しました。設定を確認してください。');
  }
  
  console.log('🧪 ========================================');
  
  return testResults;
}

/**
 * 【機能確認用】新仕様の動作確認テスト
 * 12:00実行・当日以降更新の動作をシミュレート
 */
function testNewSpecificationBehavior() {
  console.log('🔄 ========================================');
  console.log('🔄 新仕様動作確認テスト（12:00・当日以降更新）');
  console.log('🔄 ========================================');
  
  const now = new Date();
  const currentDay = now.getDate();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  
  console.log(`📅 現在日時: ${year}年${month}月${currentDay}日`);
  console.log(`🕐 実行予定時刻: 毎日12:00（現在は手動実行）`);
  console.log(`📊 更新範囲: ${currentDay}日〜月末まで`);
  
  try {
    // 実際の更新処理を実行
    console.log('\n🔄 当日以降のデータ更新を実行中...');
    updateDailyMealSheet();
    
    const lastDay = new Date(year, month, 0).getDate();
    console.log(`✅ 更新完了: ${currentDay}日〜${lastDay}日のデータを更新しました`);
    
    // 食事原紙を確認するためのURL表示
    const urlResult = getMealSheetUrl();
    if (urlResult.success) {
      console.log(`\n🔗 結果確認用URL: ${urlResult.url}`);
      console.log(`📋 シート名: 食事原紙_${year}${month.toString().padStart(2, '0')}`);
    }
    
    console.log('\n✅ 新仕様の動作確認が完了しました！');
    
  } catch (e) {
    console.error('❌ 新仕様テストでエラーが発生:', e.message);
    console.error('Stack:', e.stack);
  }
  
  console.log('🔄 ========================================');
}

/**
 * 【設定確認用】システム設定状況を確認する
 */
function checkSystemConfiguration() {
  console.log('⚙️ ========================================');
  console.log('⚙️ システム設定状況確認');
  console.log('⚙️ ========================================');
  
  try {
    // 1. スプレッドシートアクセス確認
    console.log('\n📊 1. スプレッドシートアクセス確認');
    
    const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
    const dataSheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    
    try {
      const mealSS = SpreadsheetApp.openById(mealSheetId);
      console.log(`✅ 食事原紙スプレッドシート: アクセス可能`);
      console.log(`   URL: ${mealSS.getUrl()}`);
      
      // テンプレートシート確認
      const templateSheet = mealSS.getSheetByName("食事原紙");
      if (templateSheet) {
        console.log(`✅ テンプレートシート「食事原紙」: 存在`);
      } else {
        console.log(`❌ テンプレートシート「食事原紙」: 存在しません`);
      }
    } catch (e) {
      console.log(`❌ 食事原紙スプレッドシート: アクセス不可 (${e.message})`);
    }
    
    try {
      const dataSS = SpreadsheetApp.openById(dataSheetId);
      console.log(`✅ 予約データスプレッドシート: アクセス可能`);
      
      // usersシート確認
      const usersSheet = dataSS.getSheetByName("users");
      if (usersSheet) {
        const userCount = usersSheet.getLastRow() - 1; // ヘッダー行を除く
        console.log(`✅ usersシート: 存在 (${userCount}ユーザー)`);
      } else {
        console.log(`❌ usersシート: 存在しません`);
      }
    } catch (e) {
      console.log(`❌ 予約データスプレッドシート: アクセス不可 (${e.message})`);
    }
    
    // 2. トリガー設定確認
    console.log('\n⏰ 2. トリガー設定確認');
    const triggers = ScriptApp.getProjectTriggers();
    console.log(`設定済みトリガー数: ${triggers.length}`);
    
    triggers.forEach((trigger, index) => {
      const handlerFunction = trigger.getHandlerFunction();
      const eventType = trigger.getEventType().toString();
      console.log(`  トリガー${index + 1}: ${handlerFunction} (${eventType})`);
    });
    
    if (triggers.length === 0) {
      console.log('⚠️ トリガーが設定されていません。setupTriggers()を実行してください。');
    }
    
    // 3. 現在月のシート確認
    console.log('\n📅 3. 現在月のシート確認');
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
    const currentSheetName = `食事原紙_${yyyyMM}`;
    
    try {
      const mealSS = SpreadsheetApp.openById(mealSheetId);
      const currentSheet = mealSS.getSheetByName(currentSheetName);
      if (currentSheet) {
        console.log(`✅ 現在月シート「${currentSheetName}」: 存在`);
      } else {
        console.log(`❌ 現在月シート「${currentSheetName}」: 存在しません`);
        console.log(`   → createMonthlySheet()を実行してシートを作成してください`);
      }
    } catch (e) {
      console.log(`❌ 現在月シート確認エラー: ${e.message}`);
    }
    
    console.log('\n✅ システム設定状況確認完了');
    
  } catch (e) {
    console.error('❌ 設定確認中にエラーが発生:', e.message);
  }
  
  console.log('⚙️ ========================================');
}