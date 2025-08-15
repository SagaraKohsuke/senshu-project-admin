function test(){
  const hoge = getMonthlyReservationCounts(2025, 4);
  Logger.log(hoge);
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
  
  return {
    success: true,
    year: year,
    month: month,
    breakfast: breakfastReservations,
    dinner: dinnerReservations
  };
}