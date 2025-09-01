// ==========================================
// メニューとカロリー管理機能
// ==========================================

/**
 * メニューリスト取得（内部実装）
 */
function getMenuListsImpl() {
  try {
    console.log('=== getMenuListsImpl開始 ===');
    
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
    console.error('❌ getMenuListsImpl エラー:', error);
    return {
      success: false,
      message: 'メニューリストの取得に失敗しました: ' + error.message,
      breakfast: [],
      dinner: []
    };
  }
}

/**
 * メニュー更新（内部実装）
 */
function updateMenuForCalendarImpl(calendarId, mealType, menuName, calorieValue, year, month) {
  try {
    console.log('=== updateMenuForCalendarImpl開始 ===');
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
    console.error('❌ updateMenuForCalendarImpl エラー:', error);
    return {
      success: false,
      message: 'メニュー更新中にエラーが発生しました: ' + error.message
    };
  }
}

// 既存の関数（互換性のため残す）
function updateMenuForCalendar(calendarId, mealType, menuName, calorie, year, month) {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);
  
  // シート名を生成
  const yyyyMM = `${year}${month.toString().padStart(2, "0")}`;
  const prefix = mealType === "breakfast" ? "b" : "d";
  const calendarSheetName = `${prefix}_calendar_${yyyyMM}`;
  
  // メニューシート名
  const menuSheetName = mealType === "breakfast" ? "b_menus" : "d_menus";
  
  // シートの存在確認
  const calendarSheet = ss.getSheetByName(calendarSheetName);
  let menuSheet = ss.getSheetByName(menuSheetName);
  
  if (!calendarSheet) {
    return {
      success: false,
      message: `カレンダーシート ${calendarSheetName} が見つかりません。`
    };
  }
  
  // メニューシートがなければ作成
  if (!menuSheet) {
    if (mealType === "breakfast") {
      // 朝食メニューシート作成
      const newMenuSheet = ss.insertSheet(menuSheetName);
      newMenuSheet.getRange("A1:C1").setValues([["b_menu_id", "breakfast_menu", "calorie"]]);
      newMenuSheet.getRange("A1:C1").setFontWeight("bold");
      newMenuSheet.autoResizeColumns(1, 3);
    } else {
      // 夕食メニューシート作成
      const newMenuSheet = ss.insertSheet(menuSheetName);
      newMenuSheet.getRange("A1:C1").setValues([["d_menu_id", "dinner_menu", "calorie"]]);
      newMenuSheet.getRange("A1:C1").setFontWeight("bold");
      newMenuSheet.autoResizeColumns(1, 3);
    }
    
    menuSheet = ss.getSheetByName(menuSheetName);
  }
  
  // カレンダーデータを取得
  const calendarData = calendarSheet.getDataRange().getValues();
  const menuData = menuSheet.getDataRange().getValues();
  
  // ヘッダー行の列インデックスを取得
  const calendarHeaders = calendarData[0];
  const menuHeaders = menuData[0];
  
  const calendarIdIndex = calendarHeaders.indexOf(`${prefix}_calendar_id`);
  const menuIdIndex = calendarHeaders.indexOf(`${prefix}_menu_id`);
  
  if (calendarIdIndex === -1 || menuIdIndex === -1) {
    return {
      success: false,
      message: "必要なカラムがカレンダーシートに見つかりません。"
    };
  }
  
  // メニューテーブルのインデックスを取得
  const menuIdColIndex = menuHeaders.indexOf(`${prefix}_menu_id`);
  const menuNameColIndex = mealType === "breakfast" 
    ? menuHeaders.indexOf("breakfast_menu") 
    : menuHeaders.indexOf("dinner_menu");
  const calorieColIndex = menuHeaders.indexOf("calorie");
  
  if (menuIdColIndex === -1 || menuNameColIndex === -1) {
    return {
      success: false,
      message: "必要なカラムがメニューシートに見つかりません。"
    };
  }
  
  // メニューデータを検索
  let menuId = null;
  let isNewMenu = true;
  
  // 入力されたメニュー名が既存のメニューテーブルに存在するか確認
  for (let i = 1; i < menuData.length; i++) {
    if (menuData[i][menuNameColIndex] === menuName) {
      menuId = menuData[i][menuIdColIndex];
      isNewMenu = false;
      
      // 既存メニューのカロリー情報を更新（カロリーカラムが存在する場合）
      if (calorieColIndex !== -1 && calorie !== undefined && calorie !== null) {
        menuSheet.getRange(i + 1, calorieColIndex + 1).setValue(calorie);
      }
      break;
    }
  }
  
  // 新しいメニューの場合、メニューテーブルに追加
  if (isNewMenu) {
    let newMenuId = 1;
    
    // 既存のメニューIDの最大値を取得
    if (menuData.length > 1) {
      // 念のため数値型に変換してから最大値を求める
      const menuIds = menuData.slice(1).map(row => {
        const id = row[menuIdColIndex];
        return typeof id === 'number' ? id : Number(id);
      }).filter(id => !isNaN(id));
      
      if (menuIds.length > 0) {
        newMenuId = Math.max(...menuIds) + 1;
      }
    }
    
    // 新しいメニューを追加
    const newMenuRow = [newMenuId, menuName];
    if (calorieColIndex !== -1) {
      newMenuRow.push(calorie || 0); // カロリーが未入力の場合は0
    }
    menuSheet.appendRow(newMenuRow);
    menuId = newMenuId;
  }
  
  // カレンダーテーブルのメニューIDを更新
  let rowIndex = -1;
  
  // 指定されたカレンダーIDの行を検索
  for (let i = 1; i < calendarData.length; i++) {
    if (calendarData[i][calendarIdIndex] == calendarId) {
      rowIndex = i;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return {
      success: false,
      message: `カレンダーID ${calendarId} が見つかりません。`
    };
  }
  
  // メニューIDを更新
  calendarSheet.getRange(rowIndex + 1, menuIdIndex + 1).setValue(menuId);
  
  return {
    success: true,
    message: "メニューを更新しました。",
    menuId: menuId,
    isNewMenu: isNewMenu
  };
}

/**
 * 朝食と夕食のメニュー一覧を取得する
 * @return {Object} 朝食・夕食のメニュー一覧
 */
function getMenuLists() {
  const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
  const ss = SpreadsheetApp.openById(spreadsheetId);
  
  // シートの取得
  const bMenuSheet = ss.getSheetByName("b_menus");
  const dMenuSheet = ss.getSheetByName("d_menus");
  
  // メニューデータ格納用配列
  const breakfastMenus = [];
  const dinnerMenus = [];
  
  // 朝食メニューデータの取得（シートが存在する場合）
  if (bMenuSheet) {
    const bMenuData = bMenuSheet.getDataRange().getValues();
    
    if (bMenuData.length > 1) {
      const bMenuNameIndex = bMenuData[0].indexOf("breakfast_menu");
      const bCalorieIndex = bMenuData[0].indexOf("calorie");
      
      if (bMenuNameIndex !== -1) {
        for (let i = 1; i < bMenuData.length; i++) {
          const menuName = bMenuData[i][bMenuNameIndex];
          const calorie = bCalorieIndex !== -1 ? bMenuData[i][bCalorieIndex] : 0;
          if (menuName && menuName !== "未設定") {
            breakfastMenus.push({
              name: menuName,
              calorie: calorie || 0
            });
          }
        }
      }
    }
  }
  
  // 夕食メニューデータの取得（シートが存在する場合）
  if (dMenuSheet) {
    const dMenuData = dMenuSheet.getDataRange().getValues();
    
    if (dMenuData.length > 1) {
      const dMenuNameIndex = dMenuData[0].indexOf("dinner_menu");
      const dCalorieIndex = dMenuData[0].indexOf("calorie");
      
      if (dMenuNameIndex !== -1) {
        for (let i = 1; i < dMenuData.length; i++) {
          const menuName = dMenuData[i][dMenuNameIndex];
          const calorie = dCalorieIndex !== -1 ? dMenuData[i][dCalorieIndex] : 0;
          if (menuName && menuName !== "未設定") {
            dinnerMenus.push({
              name: menuName,
              calorie: calorie || 0
            });
          }
        }
      }
    }
  }
  
  // 重複を削除してメニューを並べ替え
  const uniqueBreakfastMenus = breakfastMenus.reduce((unique, menu) => {
    const found = unique.find(m => m.name === menu.name);
    if (!found) {
      unique.push(menu);
    }
    return unique;
  }, []).sort((a, b) => a.name.localeCompare(b.name));
  
  const uniqueDinnerMenus = dinnerMenus.reduce((unique, menu) => {
    const found = unique.find(m => m.name === menu.name);
    if (!found) {
      unique.push(menu);
    }
    return unique;
  }, []).sort((a, b) => a.name.localeCompare(b.name));
  
  return {
    success: true,
    breakfast: uniqueBreakfastMenus,
    dinner: uniqueDinnerMenus
  };
}