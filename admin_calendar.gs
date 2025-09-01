function debugTest(){
  console.log('=== テスト開始 ===');
  try {
    // まずスプレッドシートの整合性をチェック
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    console.log('✅ スプレッドシート接続成功');
    
    const yyyyMM = "202508";
    console.log('検索対象:', yyyyMM);
    
    const bCalendarSheetName = "b_calendar_" + yyyyMM;
    const dCalendarSheetName = "d_calendar_" + yyyyMM;
    const bReservationSheetName = "b_reservations_" + yyyyMM;
    const dReservationSheetName = "d_reservations_" + yyyyMM;
    
    const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
    const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
    const bReservationSheet = ss.getSheetByName(bReservationSheetName);
    const dReservationSheet = ss.getSheetByName(dReservationSheetName);
    
    console.log('シート存在確認:');
    console.log(`- ${bCalendarSheetName}: ${bCalendarSheet ? 'EXISTS' : 'NOT FOUND'}`);
    console.log(`- ${dCalendarSheetName}: ${dCalendarSheet ? 'EXISTS' : 'NOT FOUND'}`);
    console.log(`- ${bReservationSheetName}: ${bReservationSheet ? 'EXISTS' : 'NOT FOUND'}`);
    console.log(`- ${dReservationSheetName}: ${dReservationSheet ? 'EXISTS' : 'NOT FOUND'}`);
    
    if (bCalendarSheet) {
      const headers = bCalendarSheet.getRange(1, 1, 1, bCalendarSheet.getLastColumn()).getValues()[0];
      console.log(`${bCalendarSheetName} ヘッダー:`, headers);
    }
    
    // メイン関数のテスト
    const result = getMonthlyReservationCounts(2025, 8);
    console.log('テスト結果:', result);
    if (result.success) {
      console.log('✅ 正常に動作しました');
      console.log(`朝食データ数: ${result.breakfast.length}`);
      console.log(`夕食データ数: ${result.dinner.length}`);
    } else {
      console.log('❌ エラー:', result.message);
    }
  } catch (e) {
    console.log('❌ 例外発生:', e.message);
    console.log('スタック:', e.stack);
  }
  console.log('=== テスト終了 ===');
}

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
 * 食事原紙を生成・更新する
 * @param {number} year 年
 * @param {number} month 月
 * @return {Object} 結果
 */
function generateMealSheet(year, month) {
  try {
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    const sheetName = "meal_sheet_" + yyyyMM;
    
    // 既存シートを削除して新規作成
    const existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    const mealSheet = ss.insertSheet(sheetName);
    
    // ユーザーデータを取得
    const usersSheet = ss.getSheetByName("users");
    if (!usersSheet) {
      return {
        success: false,
        message: "ユーザーシートが見つかりません。"
      };
    }
    
    const usersData = usersSheet.getDataRange().getValues();
    const usersHeaders = usersData[0];
    const userIdIndex = usersHeaders.indexOf("user_id");
    const userNameIndex = usersHeaders.indexOf("name");
    
    if (userIdIndex === -1 || userNameIndex === -1) {
      return {
        success: false,
        message: "ユーザーシートに必要なカラムが見つかりません。"
      };
    }
    
    // 月の日数を取得
    const daysInMonth = new Date(year, month, 0).getDate();
    
    // ヘッダー行を作成
    const headers = ["部屋番号", "名前"];
    
    // 日付ヘッダーを追加（朝食・夕食）
    for (let day = 1; day <= daysInMonth; day++) {
      const date = new Date(year, month - 1, day);
      const dayOfWeek = date.getDay();
      
      headers.push(day + "朝");
      if (dayOfWeek !== 6) { // 土曜日でなければ夕食も追加
        headers.push(day + "夕");
      }
    }
    
    // ヘッダーを設定
    mealSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    mealSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    
    // 予約データを取得
    const reservationData = getMonthlyReservationCounts(year, month);
    if (!reservationData.success) {
      return {
        success: false,
        message: "予約データの取得に失敗しました: " + reservationData.message
      };
    }
    
    // 日付別予約マップを作成
    const breakfastMap = {};
    const dinnerMap = {};
    
    for (const item of reservationData.breakfast) {
      if (item.users && Array.isArray(item.users)) {
        for (const user of item.users) {
          const dateKey = item.date;
          if (!breakfastMap[dateKey]) {
            breakfastMap[dateKey] = {};
          }
          breakfastMap[dateKey][user.userId] = true;
        }
      }
    }
    
    for (const item of reservationData.dinner) {
      if (item.users && Array.isArray(item.users)) {
        for (const user of item.users) {
          const dateKey = item.date;
          if (!dinnerMap[dateKey]) {
            dinnerMap[dateKey] = {};
          }
          dinnerMap[dateKey][user.userId] = true;
        }
      }
    }
    
    // ユーザー行を作成
    const rows = [];
    for (let i = 1; i < usersData.length; i++) {
      const userId = usersData[i][userIdIndex];
      const userName = usersData[i][userNameIndex];
      
      if (userId && userName) {
        const row = [userId, userName];
        
        // 各日の予約状況を追加
        for (let day = 1; day <= daysInMonth; day++) {
          const dateStr = year + "-" + (month < 10 ? "0" + month : month) + "-" + (day < 10 ? "0" + day : day);
          const date = new Date(year, month - 1, day);
          const dayOfWeek = date.getDay();
          
          // 朝食
          const hasBreakfast = breakfastMap[dateStr] && breakfastMap[dateStr][userId];
          row.push(hasBreakfast ? 1 : "");
          
          // 夕食（土曜日以外）
          if (dayOfWeek !== 6) {
            const hasDinner = dinnerMap[dateStr] && dinnerMap[dateStr][userId];
            row.push(hasDinner ? 1 : "");
          }
        }
        
        rows.push(row);
      }
    }
    
    // データを設定
    if (rows.length > 0) {
      mealSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    // 列幅を調整
    mealSheet.autoResizeColumns(1, headers.length);
    
    // 枠線を追加
    const totalRows = rows.length + 1;
    mealSheet.getRange(1, 1, totalRows, headers.length).setBorder(true, true, true, true, true, true);
    
    return {
      success: true,
      message: "食事原紙を生成しました。",
      sheetName: sheetName,
      url: ss.getUrl() + "#gid=" + mealSheet.getSheetId()
    };
    
  } catch (e) {
    console.error('generateMealSheet Error: ' + e.message);
    return {
      success: false,
      message: "食事原紙の生成中にエラーが発生しました: " + e.message
    };
  }
}

/**
 * 毎月01日00:00に実行される食事原紙生成処理
 * 指定されたテンプレートスプレッドシートに月次シートを作成
 * @return {Object} 結果
 */
function monthlyMealSheetGeneration() {
  try {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    
    console.log('=== 月次食事原紙生成開始 ===');
    console.log('対象月:', year + '年' + month + '月');
    
    const result = generateMonthlyMealSheet(year, month);
    
    if (result.success) {
      console.log('✅ 月次食事原紙生成成功');
      console.log('シート名:', result.sheetName);
    } else {
      console.log('❌ 月次食事原紙生成失敗:', result.message);
    }
    
    return result;
    
  } catch (e) {
    console.error('monthlyMealSheetGeneration Error: ' + e.message);
    return {
      success: false,
      message: '月次食事原紙生成中にエラーが発生しました: ' + e.message
    };
  }
}

/**
 * 指定月の食事原紙を生成する（食事原紙_yyyymm形式）
 * @param {number} year 年
 * @param {number} month 月
 * @return {Object} 結果
 */
function generateMonthlyMealSheet(year, month) {
  try {
    const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
    const dataSheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    
    const mealSs = SpreadsheetApp.openById(mealSheetId);
    const dataSs = SpreadsheetApp.openById(dataSheetId);
    
    console.log('食事原紙生成開始:', {
      year: year,
      month: month
    });
    
    // 対象月のシート名を決定（食事原紙_yyyymm形式）
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    const mealSheetName = "食事原紙_" + yyyyMM;
    
    console.log('対象シート:', mealSheetName);
    
    // 既存シートがある場合は削除
    const existingSheet = mealSs.getSheetByName(mealSheetName);
    if (existingSheet) {
      mealSs.deleteSheet(existingSheet);
      console.log('既存シートを削除しました:', mealSheetName);
    }
    
    // テンプレートシートをコピーして新しいシートを作成
    const templateSheet = mealSs.getSheetByName("食事原紙");
    if (!templateSheet) {
      return {
        success: false,
        message: "食事原紙テンプレートシートが見つかりません。"
      };
    }
    
    // テンプレートシートをコピー
    const newSheet = templateSheet.copyTo(mealSs);
    newSheet.setName(mealSheetName);
    
    console.log('テンプレートシートをコピーしました:', mealSheetName);
    
    // 月の日数を取得
    const daysInMonth = new Date(year, month, 0).getDate();
    const dayOfWeekNames = ['日', '月', '火', '水', '木', '金', '土'];
    
    // 1. タイトル行の年月を更新（A1セル）
    const titleCell = newSheet.getRange(1, 1);
    titleCell.setValue(year + "年" + month + "月度食事申し込み表　前半");
    
    // 2. 前半部分（1-16日）のヘッダー更新
    // 行2: 日付番号、行3: 曜日
    for (let day = 1; day <= Math.min(16, daysInMonth); day++) {
      const date = new Date(year, month - 1, day);
      const dayOfWeek = dayOfWeekNames[date.getDay()];
      
      // 日付と曜日の列位置: 1日→C列(3),D列(4), 2日→E列(5),F列(6)...
      const dayCol = 3 + (day - 1) * 2; // 日付列
      const dayNameCol = dayCol + 1; // 曜日列
      
      newSheet.getRange(2, dayCol).setValue(day);
      newSheet.getRange(2, dayNameCol).setValue(dayOfWeek);
    }
    
    // 3. 後半部分のタイトル更新（行36付近）
    newSheet.getRange(36, 1).setValue(year + "年" + month + "月度食事申し込み表　後半");
    
    // 4. 後半部分（17-31日）のヘッダー更新（行38）
    const backHeaderRow = 38;
    for (let day = 17; day <= daysInMonth; day++) {
      const date = new Date(year, month - 1, day);
      const dayOfWeek = dayOfWeekNames[date.getDay()];
      
      // 17日→C列, 18日→E列...
      const dayCol = 3 + (day - 17) * 2;
      const dayNameCol = dayCol + 1;
      
      newSheet.getRange(backHeaderRow, dayCol).setValue(day);
      newSheet.getRange(backHeaderRow, dayNameCol).setValue(dayOfWeek);
    }
    
    // 5. 予約データのクリア（朝食・夕食の数値セルのみ）
    // 前半部分のデータクリア（行5-35、列C以降）
    for (let row = 5; row <= 35; row++) {
      for (let day = 1; day <= Math.min(16, daysInMonth); day++) {
        const date = new Date(year, month - 1, day);
        const breakfastCol = 3 + (day - 1) * 2; // 朝食列
        const dinnerCol = breakfastCol + 1; // 夕食列
        
        // 朝食セルクリア
        const breakfastCell = newSheet.getRange(row, breakfastCol);
        const breakfastValue = breakfastCell.getValue();
        if (typeof breakfastValue === 'number' || breakfastValue === 1) {
          breakfastCell.setValue('');
        }
        
        // 夕食セル（土曜日以外）クリア
        if (date.getDay() !== 6) { // 土曜日でない場合
          const dinnerCell = newSheet.getRange(row, dinnerCol);
          const dinnerValue = dinnerCell.getValue();
          if (typeof dinnerValue === 'number' || dinnerValue === 1) {
            dinnerCell.setValue('');
          }
        }
      }
    }
    
    // 後半部分のデータクリア（行40以降、列C以降）
    for (let row = 40; row <= 75; row++) {
      for (let day = 17; day <= daysInMonth; day++) {
        const date = new Date(year, month - 1, day);
        const breakfastCol = 3 + (day - 17) * 2; // 朝食列
        const dinnerCol = breakfastCol + 1; // 夕食列
        
        // 朝食セルクリア
        const breakfastCell = newSheet.getRange(row, breakfastCol);
        const breakfastValue = breakfastCell.getValue();
        if (typeof breakfastValue === 'number' || breakfastValue === 1) {
          breakfastCell.setValue('');
        }
        
        // 夕食セル（土曜日以外）クリア
        if (date.getDay() !== 6) { // 土曜日でない場合
          const dinnerCell = newSheet.getRange(row, dinnerCol);
          const dinnerValue = dinnerCell.getValue();
          if (typeof dinnerValue === 'number' || dinnerValue === 1) {
            dinnerCell.setValue('');
          }
        }
      }
    }
    
    return {
      success: true,
      message: year + "年" + month + "月の食事原紙「" + mealSheetName + "」を作成しました。",
      sheetName: mealSheetName,
      url: mealSs.getUrl() + "#gid=" + newSheet.getSheetId()
    };
    
  } catch (e) {
    console.error('generateMonthlyMealSheet Error: ' + e.message);
    return {
      success: false,
      message: "食事原紙の生成中にエラーが発生しました: " + e.message
    };
  }
}

/**
 * テスト用：当日の朝食・夕食の記録を手動で作成する
 * 本来は18:00に自動実行される処理
 * @return {Object} 結果
 */
function testCreateDailyMealRecord() {
  try {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    const day = today.getDate();
    
    console.log('=== テスト用 当日記録作成開始 ===');
    console.log('対象日:', year + '年' + month + '月' + day + '日');
    
    const result = createDailyMealRecord(year, month, day);
    
    if (result.success) {
      console.log('✅ 当日記録作成成功');
      console.log('作成されたレコード:', result.records);
    } else {
      console.log('❌ 当日記録作成失敗:', result.message);
    }
    
    return result;
    
  } catch (e) {
    console.error('testCreateDailyMealRecord Error: ' + e.message);
    return {
      success: false,
      message: 'テスト実行中にエラーが発生しました: ' + e.message
    };
  }
}

/**
 * 指定日の朝食・夕食の記録を食事原紙に書き込む（最適化版）
 * 実行時間に基づいて適切な「食事原紙_yyyymm」シートを選択
 * @param {number} year 年
 * @param {number} month 月
 * @param {number} day 日
 * @return {Object} 結果
 */
function createDailyMealRecord(year, month, day) {
  try {
    const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
    const dataSheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    
    const mealSs = SpreadsheetApp.openById(mealSheetId);
    const dataSs = SpreadsheetApp.openById(dataSheetId);
    
    const dateStr = year + "-" + (month < 10 ? "0" + month : month) + "-" + (day < 10 ? "0" + day : day);
    const targetDate = new Date(year, month - 1, day);
    const dayOfWeek = targetDate.getDay();
    
    // 実行時間を取得
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    
    console.log('食事原紙記録作成対象:', {
      dateStr: dateStr,
      day: day,
      dayOfWeek: dayOfWeek,
      executionTime: now.toISOString(),
      targetYear: year,
      targetMonth: month,
      currentYear: currentYear,
      currentMonth: currentMonth
    });
    
    // 適切なシート名を決定（食事原紙_yyyymm形式）
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    const mealSheetName = "食事原紙_" + yyyyMM;
    
    // 食事原紙シートを取得
    const mealSheet = mealSs.getSheetByName(mealSheetName);
    if (!mealSheet) {
      return {
        success: false,
        message: "食事原紙シート「" + mealSheetName + "」が見つかりません。"
      };
    }
    
    // 予約データを取得
    const reservationData = getMonthlyReservationCounts(year, month);
    if (!reservationData.success) {
      return {
        success: false,
        message: "予約データの取得に失敗しました: " + reservationData.message
      };
    }
    
    // CSVテンプレート構造に基づく列位置計算
    let breakfastCol = -1;
    let dinnerCol = -1;
    let targetRowStart = -1;
    
    if (day <= 16) {
      // 前半部分（1-16日）
      breakfastCol = 3 + (day - 1) * 2; // 朝食列: 1日→C(3), 2日→E(5)...
      dinnerCol = breakfastCol + 1; // 夕食列: 1日→D(4), 2日→F(6)...
      targetRowStart = 5; // 前半部分の開始行
    } else {
      // 後半部分（17-31日）
      breakfastCol = 3 + (day - 17) * 2; // 朝食列: 17日→C(3), 18日→E(5)...
      dinnerCol = breakfastCol + 1; // 夕食列: 17日→D(4), 18日→F(6)...
      targetRowStart = 40; // 後半部分の開始行
    }
    
    console.log('列位置計算結果:', {
      day: day,
      breakfastCol: breakfastCol,
      dinnerCol: dinnerCol,
      targetRowStart: targetRowStart,
      isSaturday: dayOfWeek === 6,
      targetSheet: mealSheetName
    });
    
    // ユーザー行マッピングを作成（A列の部屋番号を基準）
    const userRows = {};
    const maxRow = targetRowStart + 30; // 想定される最大行数
    
    for (let row = targetRowStart; row <= maxRow; row++) {
      const userId = mealSheet.getRange(row, 1).getValue();
      if (userId && typeof userId !== 'undefined' && userId !== '') {
        userRows[userId.toString()] = row;
      }
    }
    
    console.log('ユーザー行マッピング作成完了:', {
      userCount: Object.keys(userRows).length,
      userIds: Object.keys(userRows).slice(0, 5) // 最初の5つのユーザーIDを表示
    });
    
    let recordsCreated = 0;
    const createdRecords = [];
    
    // 朝食の記録を作成
    const breakfastData = reservationData.breakfast.find(item => item.date === dateStr);
    if (breakfastData && breakfastData.users && Array.isArray(breakfastData.users)) {
      for (const user of breakfastData.users) {
        const userRow = userRows[user.userId.toString()];
        if (userRow) {
          mealSheet.getRange(userRow, breakfastCol).setValue(1);
          recordsCreated++;
          createdRecords.push({
            userId: user.userId,
            userName: user.userName,
            mealType: 'breakfast',
            row: userRow,
            col: breakfastCol
          });
        } else {
          console.log('朝食：ユーザー行が見つかりません:', user.userId);
        }
      }
    }
    
    // 夕食の記録を作成（土曜日以外）
    if (dayOfWeek !== 6) {
      const dinnerData = reservationData.dinner.find(item => item.date === dateStr);
      if (dinnerData && dinnerData.users && Array.isArray(dinnerData.users)) {
        for (const user of dinnerData.users) {
          const userRow = userRows[user.userId.toString()];
          if (userRow) {
            mealSheet.getRange(userRow, dinnerCol).setValue(1);
            recordsCreated++;
            createdRecords.push({
              userId: user.userId,
              userName: user.userName,
              mealType: 'dinner',
              row: userRow,
              col: dinnerCol
            });
          } else {
            console.log('夕食：ユーザー行が見つかりません:', user.userId);
          }
        }
      }
    }
    
    console.log('食事原紙記録作成完了:', {
      date: dateStr,
      targetSheet: mealSheetName,
      totalRecords: recordsCreated,
      breakfastCount: breakfastData ? breakfastData.users.length : 0,
      dinnerCount: (dayOfWeek !== 6 && reservationData.dinner.find(item => item.date === dateStr)) ? 
        reservationData.dinner.find(item => item.date === dateStr).users.length : 0,
      recordDetails: createdRecords
    });
    
    return {
      success: true,
      message: "食事原紙「" + mealSheetName + "」に記録を作成しました。",
      date: dateStr,
      day: day,
      sheetName: mealSheetName,
      recordsCreated: recordsCreated,
      records: createdRecords
    };
    
  } catch (e) {
    console.error('createDailyMealRecord Error: ' + e.message);
    return {
      success: false,
      message: "食事記録作成中にエラーが発生しました: " + e.message
    };
  }
}

/**
 * 指定した日の募集状態を切り替える（is_activeフィールドを使用）
 * @param {string} date 対象日付 (YYYY-MM-DD形式)
 * @param {string} mealType 食事タイプ ("breakfast" または "dinner")
 * @param {number} year 年
 * @param {number} month 月
 * @return {Object} 結果
 */
function toggleRecruitmentStop(date, mealType, year, month) {
  try {
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
    const prefix = mealType === "breakfast" ? "b" : "d";
    const calendarSheetName = `${prefix}_calendar_${yyyyMM}`;
    
    const calendarSheet = ss.getSheetByName(calendarSheetName);
    if (!calendarSheet) {
      return {
        success: false,
        message: `カレンダーシート ${calendarSheetName} が見つかりません。`
      };
    }
    
    const calendarData = calendarSheet.getDataRange().getValues();
    const headers = calendarData[0];
    
    const calendarIdIndex = headers.indexOf(`${prefix}_calendar_id`);
    const dateIndex = headers.indexOf("date");
    const isActiveIndex = headers.indexOf("is_active");
    
    if (calendarIdIndex === -1 || dateIndex === -1 || isActiveIndex === -1) {
      return {
        success: false,
        message: "必要なカラムが見つかりません。"
      };
    }
    
    // 対象日付の行を検索
    let targetRowIndex = -1;
    for (let i = 1; i < calendarData.length; i++) {
      const rowDate = calendarData[i][dateIndex];
      let dateStr;
      
      if (rowDate instanceof Date) {
        dateStr = formatDate(rowDate);
      } else {
        dateStr = rowDate;
      }
      
      if (dateStr === date) {
        targetRowIndex = i;
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      return {
        success: false,
        message: "指定された日付が見つかりません。"
      };
    }
    
    // is_activeを切り替え
    const currentActive = calendarData[targetRowIndex][isActiveIndex];
    const newActive = !currentActive;
    
    calendarSheet.getRange(targetRowIndex + 1, isActiveIndex + 1).setValue(newActive);
    
    return {
      success: true,
      isActive: newActive,
      message: newActive ? "休みを解除しました" : "休みにしました"
    };
    
  } catch (e) {
    console.error('toggleRecruitmentStop Error: ' + e.message);
    return {
      success: false,
      message: "募集状態の変更中にエラーが発生しました: " + e.message
    };
  }
}

/**
 * 募集停止状況を取得する（is_activeフィールドを使用）
 * @param {number} year 年
 * @param {number} month 月
 * @return {Object} 募集停止情報
 */
function getRecruitmentStops(year, month) {
  try {
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
    const bCalendarSheetName = `b_calendar_${yyyyMM}`;
    const dCalendarSheetName = `d_calendar_${yyyyMM}`;
    
    const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
    const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
    
    const stops = {};
    
    // 朝食カレンダーの処理
    if (bCalendarSheet) {
      const bCalendarData = bCalendarSheet.getDataRange().getValues();
      if (bCalendarData.length > 1) {
        const headers = bCalendarData[0];
        const dateIndex = headers.indexOf("date");
        const isActiveIndex = headers.indexOf("is_active");
        
        if (dateIndex !== -1 && isActiveIndex !== -1) {
          for (let i = 1; i < bCalendarData.length; i++) {
            const rowDate = bCalendarData[i][dateIndex];
            const isActive = bCalendarData[i][isActiveIndex];
            
            let dateStr;
            if (rowDate instanceof Date) {
              dateStr = formatDate(rowDate);
            } else {
              dateStr = rowDate;
            }
            
            if (!isActive) {
              if (!stops[dateStr]) {
                stops[dateStr] = {};
              }
              stops[dateStr]['breakfast'] = true;
            }
          }
        }
      }
    }
    
    // 夕食カレンダーの処理
    if (dCalendarSheet) {
      const dCalendarData = dCalendarSheet.getDataRange().getValues();
      if (dCalendarData.length > 1) {
        const headers = dCalendarData[0];
        const dateIndex = headers.indexOf("date");
        const isActiveIndex = headers.indexOf("is_active");
        
        if (dateIndex !== -1 && isActiveIndex !== -1) {
          for (let i = 1; i < dCalendarData.length; i++) {
            const rowDate = dCalendarData[i][dateIndex];
            const isActive = dCalendarData[i][isActiveIndex];
            
            let dateStr;
            if (rowDate instanceof Date) {
              dateStr = formatDate(rowDate);
            } else {
              dateStr = rowDate;
            }
            
            if (!isActive) {
              if (!stops[dateStr]) {
                stops[dateStr] = {};
              }
              stops[dateStr]['dinner'] = true;
            }
          }
        }
      }
    }
    
    return { success: true, stops: stops };
    
  } catch (e) {
    console.error('getRecruitmentStops Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ==========================================
// 食事原紙作成機能
// ==========================================

/**
 * 月次食事原紙シート作成（内部実装）
 */
function createMonthlyMealSheet(year, month) {
  try {
    console.log('=== createMonthlyMealSheetImpl開始 ===');
    console.log('パラメータ:', year, month);
    
    const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
    const mealSS = SpreadsheetApp.openById(mealSheetId);
    
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    const newSheetName = "食事原紙_" + yyyyMM;
    
    // 既存シートの確認
    const existingSheet = mealSS.getSheetByName(newSheetName);
    if (existingSheet) {
      console.log('既存のシートが見つかりました:', newSheetName);
      return {
        success: true,
        message: `シート「${newSheetName}」は既に存在します`,
        sheetName: newSheetName
      };
    }
    
    // テンプレートシートを取得
    const templateSheet = mealSS.getSheetByName("Template");
    if (!templateSheet) {
      return {
        success: false,
        message: 'テンプレートシートが見つかりません'
      };
    }
    
    // テンプレートをコピーして新しいシートを作成
    const newSheet = templateSheet.copyTo(mealSS);
    newSheet.setName(newSheetName);
    console.log('✅ テンプレートをコピー完了:', newSheetName);
    
    // 日付データを生成
    const daysInMonth = new Date(year, month, 0).getDate();
    const dates = [];
    for (let day = 1; day <= daysInMonth; day++) {
      dates.push(day);
    }
    
    // 日付ヘッダーを設定（2行目に日付を配置）
    if (daysInMonth > 0) {
      const dateRange = newSheet.getRange(2, 2, 1, daysInMonth);
      dateRange.setValues([dates]);
    }
    
    // 週末マーカーの適用（土日列に黄色背景 #FFFF00）
    for (let day = 1; day <= daysInMonth; day++) {
      const date = new Date(year, month - 1, day);
      const dayOfWeek = date.getDay();
      const col = day + 1; // 2列目から開始
      
      // 土曜日(6)または日曜日(0)の場合
      if (dayOfWeek === 0 || dayOfWeek === 6) {
        console.log(`週末マーカー適用: ${year}/${month}/${day} (${dayOfWeek === 0 ? '日' : '土'}曜日)`);
        
        // 前半セクション (5-37行) - 数式保護行40をスキップ
        const frontRange = newSheet.getRange(5, col, 33, 1);
        frontRange.setBackground('#FFFF00');
        
        // 後半セクション (45-77行) - 数式保護行44, 79, 80をスキップ
        const backRange = newSheet.getRange(45, col, 33, 1);
        backRange.setBackground('#FFFF00');
      }
    }
    
    // 構造的配置の実装
    // 42行目に後半セクションヘッダー「後半」を配置
    newSheet.getRange(42, 1).setValue("後半");
    
    // 数式保護行の確認（40, 44, 79, 80行目）
    console.log('✅ 数式保護行をスキップ: 40, 44, 79, 80行目');
    
    // 予約データを取得して反映
    const reservationData = getMonthlyReservationCountsImpl(year, month);
    if (reservationData.success) {
      updateMealSheetWithData(newSheet, reservationData, year, month);
    }
    
    console.log('✅ 月次食事原紙シート作成完了:', newSheetName);
    
    return {
      success: true,
      message: `食事原紙「${newSheetName}」を作成しました`,
      sheetName: newSheetName,
      url: mealSS.getUrl() + "#gid=" + newSheet.getSheetId()
    };
    
  } catch (error) {
    console.error('❌ createMonthlyMealSheetImpl エラー:', error);
    return {
      success: false,
      message: 'シート作成中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 食事原紙シートにデータを反映（内部実装）
 */
function updateMealSheetWithData(sheet, reservationData, year, month) {
  try {
    console.log('=== updateMealSheetWithDataImpl開始 ===');
    
    const daysInMonth = new Date(year, month, 0).getDate();
    
    // 朝食データの反映
    for (const breakfast of reservationData.breakfast) {
      const date = new Date(breakfast.date);
      const day = date.getDate();
      const col = day + 1; // 2列目から開始
      
      if (day >= 1 && day <= daysInMonth) {
        // 前半セクション（5-37行）に朝食データを配置
        const rowIndex = 5 + ((day - 1) % 31); // 31日サイクル（数式保護行を考慮）
        if (rowIndex <= 37 && rowIndex !== 40) { // 数式保護行40をスキップ
          sheet.getRange(rowIndex, col).setValue(breakfast.count);
          
          // ユーザー名リストを隣接セルに配置
          if (breakfast.users.length > 0) {
            const userNames = breakfast.users.map(user => user.userName).join(', ');
            // メニュー情報も追加
            const menuInfo = breakfast.menuName !== '未設定' ? 
              `${breakfast.menuName}(${breakfast.calorie}kcal)` : '';
            const cellValue = menuInfo ? `${userNames}\n${menuInfo}` : userNames;
            sheet.getRange(rowIndex + 1, col).setValue(cellValue);
          }
        }
      }
    }
    
    // 夕食データの反映
    for (const dinner of reservationData.dinner) {
      const date = new Date(dinner.date);
      const day = date.getDate();
      const col = day + 1; // 2列目から開始
      
      if (day >= 1 && day <= daysInMonth) {
        // 後半セクション（45-77行）に夕食データを配置
        const rowIndex = 45 + ((day - 1) % 31); // 31日サイクル（数式保護行を考慮）
        if (rowIndex <= 77 && rowIndex !== 79 && rowIndex !== 80) { // 数式保護行79, 80をスキップ
          sheet.getRange(rowIndex, col).setValue(dinner.count);
          
          // ユーザー名リストを隣接セルに配置
          if (dinner.users.length > 0) {
            const userNames = dinner.users.map(user => user.userName).join(', ');
            // メニュー情報も追加
            const menuInfo = dinner.menuName !== '未設定' ? 
              `${dinner.menuName}(${dinner.calorie}kcal)` : '';
            const cellValue = menuInfo ? `${userNames}\n${menuInfo}` : userNames;
            sheet.getRange(rowIndex + 1, col).setValue(cellValue);
          }
        }
      }
    }
    
    console.log('✅ 食事原紙データ反映完了');
    
  } catch (error) {
    console.error('❌ updateMealSheetWithDataImpl エラー:', error);
  }
}

function getMonthlyReservationCounts(year, month) {
  try {
    console.log('=== getMonthlyReservationCounts開始: ' + year + '年' + month + '月 ===');
    
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
      const dateStr = formatDate(date);
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
      const dateStr = formatDate(date);
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
    console.error('❌ getMonthlyReservationCounts エラー:', error);
    console.error('エラースタック:', error.stack);
    return {
      success: false,
      message: '処理中にエラーが発生しました: ' + error.message,
      breakfast: [],
      dinner: []
    };
  }
}