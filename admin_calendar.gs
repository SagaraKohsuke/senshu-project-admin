function test(){
  const hoge = getMonthlyReservationCounts(2025, 4);
  Logger.log(hoge);
}
function getMonthlyReservationCounts(year, month) { 
  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const bCalendarSheetName = `b_calendar_${yyyyMM}`;
  const dCalendarSheetName = `d_calendar_${yyyyMM}`;
  const bReservationSheetName = `b_reservations_${yyyyMM}`;
  const dReservationSheetName = `d_reservations_${yyyyMM}`;
  
  // シートの存在確認
  const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
  const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
  const bReservationSheet = ss.getSheetByName(bReservationSheetName);
  const dReservationSheet = ss.getSheetByName(dReservationSheetName);
  const usersSheet = ss.getSheetByName("users");
  const bMenuSheet = ss.getSheetByName("b_menus");
  const dMenuSheet = ss.getSheetByName("d_menus");
  
  if (!bCalendarSheet || !dCalendarSheet) {
    return {
      success: false,
      message: `カレンダーシート ${bCalendarSheetName} または ${dCalendarSheetName} が見つかりません。`
    };
  }
  
  if (!bReservationSheet || !dReservationSheet) {
    return {
      success: false,
      message: `予約シート ${bReservationSheetName} または ${dReservationSheetName} が見つかりません。`
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
  
  return {
    success: true,
    year: year,
    month: month,
    breakfast: breakfastReservations,
    dinner: dinnerReservations
  };
}