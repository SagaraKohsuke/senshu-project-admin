// ==========================================
// メニューとカロリー管理機能
// ==========================================

/**
 * メニューリスト取得
 */
function getMenuLists() {
  try {
    console.log('=== getMenuLists開始 ===');
    
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
 * メニュー更新
 */
function updateMenuForCalendar(calendarId, mealType, menuName, calorieValue, year, month) {
  try {
    console.log('=== updateMenuForCalendar開始 ===');
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
    const calendarIdColumnName2 = mealType === 'breakfast' ? 'b_calendar_id' : 'd_calendar_id';
    const menuIdColumnName2 = mealType === 'breakfast' ? 'b_menu_id' : 'd_menu_id';
    
    const calendarIdIndex = calendarHeaders.indexOf(calendarIdColumnName2);
    const calendarMenuIdIndex = calendarHeaders.indexOf(menuIdColumnName2);
    
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

