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
 * 食事原紙のスプレッドシートURLを取得する
 * @return {Object} 結果とURL
 */
function getMealSheetUrl() {
  try {
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    return {
      success: true,
      url: ss.getUrl()
    };
  } catch (e) {
    console.error('getMealSheetUrl Error: ' + e.message);
    return {
      success: false,
      message: "スプレッドシートのURL取得に失敗しました: " + e.message
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
      message: newActive ? "募集を再開しました" : "募集を停止しました"
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
  
  // 朝食メニューマップの作成
  const bMenuMap = {};
  if (bMenuData.length > 1) {
    const bMenuIdIndex = bMenuData[0].indexOf("b_menu_id");
    const bMenuNameIndex = bMenuData[0].indexOf("breakfast_menu");
    
    if (bMenuIdIndex !== -1 && bMenuNameIndex !== -1) {
      for (let i = 1; i < bMenuData.length; i++) {
        const menuId = bMenuData[i][bMenuIdIndex];
        const menuName = bMenuData[i][bMenuNameIndex];
        bMenuMap[menuId] = menuName;
      }
    }
  }
  
  // 夕食メニューマップの作成
  const dMenuMap = {};
  if (dMenuData.length > 1) {
    const dMenuIdIndex = dMenuData[0].indexOf("d_menu_id");
    const dMenuNameIndex = dMenuData[0].indexOf("dinner_menu");
    
    if (dMenuIdIndex !== -1 && dMenuNameIndex !== -1) {
      for (let i = 1; i < dMenuData.length; i++) {
        const menuId = dMenuData[i][dMenuIdIndex];
        const menuName = dMenuData[i][dMenuNameIndex];
        dMenuMap[menuId] = menuName;
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
        menuName: bMenuMap[menuId] || "未設定"
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
        menuName: dMenuMap[menuId] || "未設定"
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