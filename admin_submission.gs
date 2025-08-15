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

    // 月次の予約データを取得
    const reservationData = getMonthlyReservationCounts(year, month);
    if (!reservationData.success) {
      console.error("予約データの取得に失敗しました:", reservationData.message);
      return;
    }

    // ヘッダー情報（タイトルのみ）を更新
    const titleRange = targetSheet.getRange("A1");
    const currentTitle = titleRange.getValue();
    if (currentTitle && currentTitle.toString().includes('月度')) {
      titleRange.setValue(currentTitle.toString().replace(/\d+月度/, `${month}月度`));
    }

    // 前半・後半ブロックごとにユーザーIDと行のマッピングを作成し、氏名を設定する
    const mapUsersAndSetNames = (startRow, endRow) => {
      const userRowMap = {};
      const idRange = targetSheet.getRange(`A${startRow}:A${endRow}`);
      const nameRange = targetSheet.getRange(`B${startRow}:B${endRow}`);
      const idValues = idRange.getValues();
      const namesToSet = [];

      for (let i = 0; i < idValues.length; i++) {
        const userId = idValues[i][0];
        if (userId && !isNaN(userId)) {
          userRowMap[userId] = startRow + i;
          namesToSet.push([userIdToNameMap[userId] || '']);
        } else {
          namesToSet.push(['']);
        }
      }
      nameRange.setValues(namesToSet);
      return userRowMap;
    };

    const userRowMap_1_16 = mapUsersAndSetNames(5, 37);
    const userRowMap_17_31 = mapUsersAndSetNames(45, 77);

    // 既存のデータをクリア（3列目以降）
    const maxCol = targetSheet.getMaxColumns();
    if (maxCol > 2) {
      targetSheet.getRange(5, 3, 73, maxCol - 2).clearContent();
    }

    // 予約データ（「1」）を正しいブロックのセルに書き込む
    const dataToUpdate = [];
    const { breakfast: breakfastReservations, dinner: dinnerReservations } = reservationData;

    const processReservations = (reservations, isDinner) => {
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
    };

    processReservations(breakfastReservations, false);
    processReservations(dinnerReservations, true);
    
    dataToUpdate.forEach(data => {
      targetSheet.getRange(data.row, data.col).setValue(data.value);
    });
    
    console.log(`${sheetName} の予約データを更新しました。`);

  } catch (e) {
    console.error('updateDailyMealSheet Error: ' + e.message + " Stack: " + e.stack);
  }
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