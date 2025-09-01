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
// ==========================================

// ==========================================
// フロントエンド用ラッパー関数群
// Google Apps Scriptでは、フロントエンドから直接アクセスできるのは
// admin_main.gsの関数のみなので、他のファイルの関数はここでラップする
// ==========================================

/**
 * フロントエンド用：月次予約データを取得
 * 注意：フロントエンドからは getMonthlyReservationCounts という名前で呼び出される
 */
function getMonthlyReservationCounts(year, month) {
  try {
    console.log('=== フロントエンド用getMonthlyReservationCounts開始 ===');
    console.log('パラメータ:', year, month);
    
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

/**
 * フロントエンド用：募集停止データを取得
 */
function getRecruitmentStops(year, month) {
  try {
    console.log('=== フロントエンド用getRecruitmentStops開始 ===');
    console.log('パラメータ:', year, month);
    
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // recruitment_stopsシートを取得
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
      console.log('recruitment_stopsシートにデータがありません。');
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
    
    if (dateIndex === -1 || breakfastIndex === -1 || dinnerIndex === -1 || isActiveIndex === -1) {
      console.error('recruitment_stopsシートに必要な列が見つかりません');
      return {
        success: false,
        message: 'recruitment_stopsシートの形式が正しくありません',
        stops: {}
      };
    }
    
    const stops = {};
    const targetYearMonth = year + '-' + (month < 10 ? '0' + month : month);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[dateIndex];
      const isBreakfastStopped = row[breakfastIndex];
      const isDinnerStopped = row[dinnerIndex];
      const isActive = row[isActiveIndex];
      
      if (!isActive) continue; // 無効なレコードはスキップ
      
      let dateStr;
      if (date instanceof Date) {
        dateStr = date.getFullYear() + '-' + 
                 (date.getMonth() + 1).toString().padStart(2, '0') + '-' + 
                 date.getDate().toString().padStart(2, '0');
      } else if (typeof date === 'string') {
        dateStr = date;
      } else {
        continue;
      }
      
      // 指定された年月のデータのみ処理
      if (dateStr.startsWith(targetYearMonth)) {
        stops[dateStr] = {
          breakfast: !!isBreakfastStopped,
          dinner: !!isDinnerStopped
        };
      }
    }
    
    console.log('✅ 募集停止データ取得完了:', stops);
    
    return {
      success: true,
      stops: stops
    };
    
  } catch (error) {
    console.error('❌ getRecruitmentStops エラー:', error);
    return {
      success: false,
      message: '募集停止データの取得に失敗しました: ' + error.message,
      stops: {}
    };
  }
}

/**
 * フロントエンド用：メニューリストを取得
 */
function getMenuLists() {
  try {
    console.log('=== フロントエンド用getMenuLists開始 ===');
    
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    const bMenuSheet = ss.getSheetByName("b_menus");
    const dMenuSheet = ss.getSheetByName("d_menus");
    
    let breakfastMenus = [];
    let dinnerMenus = [];
    
    // 朝食メニューの取得
    if (bMenuSheet) {
      const bMenuData = bMenuSheet.getDataRange().getValues();
      if (bMenuData.length > 1) {
        const bMenuHeaders = bMenuData[0];
        const bMenuNameIndex = bMenuHeaders.indexOf("breakfast_menu");
        const bCalorieIndex = bMenuHeaders.indexOf("calorie");
        
        if (bMenuNameIndex !== -1) {
          for (let i = 1; i < bMenuData.length; i++) {
            const menuName = bMenuData[i][bMenuNameIndex];
            const calorie = bCalorieIndex !== -1 ? (bMenuData[i][bCalorieIndex] || 0) : 0;
            
            if (menuName && menuName.trim() !== '') {
              breakfastMenus.push({
                name: menuName.trim(),
                calorie: Number(calorie) || 0
              });
            }
          }
        }
      }
    }
    
    // 夕食メニューの取得
    if (dMenuSheet) {
      const dMenuData = dMenuSheet.getDataRange().getValues();
      if (dMenuData.length > 1) {
        const dMenuHeaders = dMenuData[0];
        const dMenuNameIndex = dMenuHeaders.indexOf("dinner_menu");
        const dCalorieIndex = dMenuHeaders.indexOf("calorie");
        
        if (dMenuNameIndex !== -1) {
          for (let i = 1; i < dMenuData.length; i++) {
            const menuName = dMenuData[i][dMenuNameIndex];
            const calorie = dCalorieIndex !== -1 ? (dMenuData[i][dCalorieIndex] || 0) : 0;
            
            if (menuName && menuName.trim() !== '') {
              dinnerMenus.push({
                name: menuName.trim(),
                calorie: Number(calorie) || 0
              });
            }
          }
        }
      }
    }
    
    // 名前でソート
    breakfastMenus.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
    dinnerMenus.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
    
    console.log('✅ メニューリスト取得完了:', {
      breakfastCount: breakfastMenus.length,
      dinnerCount: dinnerMenus.length
    });
    
    return {
      success: true,
      breakfast: breakfastMenus,
      dinner: dinnerMenus
    };
    
  } catch (error) {
    console.error('❌ getMenuLists エラー:', error);
    return {
      success: false,
      message: 'メニューリストの取得に失敗しました: ' + error.message,
      breakfast: [],
      dinner: []
    };
  }
}

/**
 * フロントエンド用：食事原紙URLを取得
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
 * フロントエンド用：メニュー更新
 */
function updateMenuForCalendar(calendarId, mealType, menuName, calorieValue, year, month) {
  try {
    console.log('=== フロントエンド用updateMenuForCalendar開始 ===');
    console.log('パラメータ:', { calendarId, mealType, menuName, calorieValue, year, month });
    
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    const menuSheetName = mealType === 'breakfast' ? 'b_menus' : 'd_menus';
    const menuSheet = ss.getSheetByName(menuSheetName);
    
    if (!menuSheet) {
      return {
        success: false,
        message: `${menuSheetName}シートが見つかりません`
      };
    }
    
    const menuData = menuSheet.getDataRange().getValues();
    const headers = menuData[0];
    const menuNameColumnName = mealType === 'breakfast' ? 'breakfast_menu' : 'dinner_menu';
    const menuIdColumnName = mealType === 'breakfast' ? 'b_menu_id' : 'd_menu_id';
    
    const menuNameIndex = headers.indexOf(menuNameColumnName);
    const menuIdIndex = headers.indexOf(menuIdColumnName);
    const calorieIndex = headers.indexOf('calorie');
    
    if (menuNameIndex === -1 || menuIdIndex === -1) {
      return {
        success: false,
        message: '必要な列が見つかりません'
      };
    }
    
    // 既存のメニューを検索
    let existingMenuId = null;
    let isNewMenu = true;
    
    for (let i = 1; i < menuData.length; i++) {
      if (menuData[i][menuNameIndex] === menuName) {
        existingMenuId = menuData[i][menuIdIndex];
        isNewMenu = false;
        
        // カロリー値が異なる場合は更新
        if (calorieIndex !== -1 && menuData[i][calorieIndex] !== calorieValue) {
          menuSheet.getRange(i + 1, calorieIndex + 1).setValue(calorieValue || 0);
          console.log('既存メニューのカロリーを更新:', menuName, calorieValue);
        }
        break;
      }
    }
    
    // 新しいメニューの場合は追加
    if (isNewMenu) {
      const newMenuId = new Date().getTime(); // 簡単なID生成
      const newRow = Array(headers.length).fill('');
      newRow[menuIdIndex] = newMenuId;
      newRow[menuNameIndex] = menuName;
      if (calorieIndex !== -1) {
        newRow[calorieIndex] = calorieValue || 0;
      }
      
      menuSheet.appendRow(newRow);
      existingMenuId = newMenuId;
      console.log('新しいメニューを追加:', menuName, calorieValue);
    }
    
    // カレンダーのmenu_idを更新
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    const calendarSheetName = mealType === 'breakfast' ? `b_calendar_${yyyyMM}` : `d_calendar_${yyyyMM}`;
    const calendarSheet = ss.getSheetByName(calendarSheetName);
    
    if (!calendarSheet) {
      return {
        success: false,
        message: `${calendarSheetName}シートが見つかりません`
      };
    }
    
    const calendarData = calendarSheet.getDataRange().getValues();
    const calendarHeaders = calendarData[0];
    const calendarIdColumnName = mealType === 'breakfast' ? 'b_calendar_id' : 'd_calendar_id';
    const menuIdColumnName = mealType === 'breakfast' ? 'b_menu_id' : 'd_menu_id';
    
    const calendarIdIndex = calendarHeaders.indexOf(calendarIdColumnName);
    const calendarMenuIdIndex = calendarHeaders.indexOf(menuIdColumnName);
    
    if (calendarIdIndex === -1 || calendarMenuIdIndex === -1) {
      return {
        success: false,
        message: 'カレンダーシートに必要な列が見つかりません'
      };
    }
    
    // カレンダーのmenu_idを更新
    for (let i = 1; i < calendarData.length; i++) {
      if (calendarData[i][calendarIdIndex] == calendarId) {
        calendarSheet.getRange(i + 1, calendarMenuIdIndex + 1).setValue(existingMenuId);
        console.log('カレンダーのmenu_idを更新:', calendarId, existingMenuId);
        break;
      }
    }
    
    return {
      success: true,
      menuId: existingMenuId,
      isNewMenu: isNewMenu,
      message: isNewMenu ? 'メニューを新規追加しました' : 'メニューを更新しました'
    };
    
  } catch (error) {
    console.error('❌ updateMenuForCalendar エラー:', error);
    return {
      success: false,
      message: 'メニュー更新中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * フロントエンド用：募集停止切り替え
 */
function toggleRecruitmentStop(date, mealType, year, month) {
  try {
    console.log('=== フロントエンド用toggleRecruitmentStop開始 ===');
    console.log('パラメータ:', { date, mealType, year, month });
    
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // recruitment_stopsシートを取得または作成
    let recruitmentStopsSheet = ss.getSheetByName("recruitment_stops");
    if (!recruitmentStopsSheet) {
      recruitmentStopsSheet = ss.insertSheet("recruitment_stops");
      // ヘッダー行を作成
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
      // 既存レコードを更新
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
      // 新規レコードを作成
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
    
    console.log('✅ 募集停止状態を更新:', { date, mealType, message });
    
    return {
      success: true,
      message: message
    };
    
  } catch (error) {
    console.error('❌ toggleRecruitmentStop エラー:', error);
    return {
      success: false,
      message: '募集停止の切り替え中にエラーが発生しました: ' + error.message
    };
  }
}