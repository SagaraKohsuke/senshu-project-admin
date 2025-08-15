function updateMenuForCalendar(calendarId, mealType, menuName, year, month) {
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
      newMenuSheet.getRange("A1:B1").setValues([["b_menu_id", "breakfast_menu"]]);
      newMenuSheet.getRange("A1:B1").setFontWeight("bold");
      newMenuSheet.autoResizeColumns(1, 2);
    } else {
      // 夕食メニューシート作成
      const newMenuSheet = ss.insertSheet(menuSheetName);
      newMenuSheet.getRange("A1:B1").setValues([["d_menu_id", "dinner_menu"]]);
      newMenuSheet.getRange("A1:B1").setFontWeight("bold");
      newMenuSheet.autoResizeColumns(1, 2);
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
    menuSheet.appendRow([newMenuId, menuName]);
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
      
      if (bMenuNameIndex !== -1) {
        for (let i = 1; i < bMenuData.length; i++) {
          const menuName = bMenuData[i][bMenuNameIndex];
          if (menuName && menuName !== "未設定") {
            breakfastMenus.push(menuName);
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
      
      if (dMenuNameIndex !== -1) {
        for (let i = 1; i < dMenuData.length; i++) {
          const menuName = dMenuData[i][dMenuNameIndex];
          if (menuName && menuName !== "未設定") {
            dinnerMenus.push(menuName);
          }
        }
      }
    }
  }
  
  // 重複を削除してメニューを並べ替え
  const uniqueBreakfastMenus = [...new Set(breakfastMenus)].sort();
  const uniqueDinnerMenus = [...new Set(dinnerMenus)].sort();
  
  return {
    success: true,
    breakfast: uniqueBreakfastMenus,
    dinner: uniqueDinnerMenus
  };
}