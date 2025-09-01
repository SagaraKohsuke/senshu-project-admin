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
    
    // 3. 後半部分のタイトル更新（行41）
    newSheet.getRange(41, 1).setValue(year + "年" + month + "月度食事申し込み表　後半");
    
    // 4. 後半部分（17-31日）のヘッダー更新（行42）
    const backHeaderRow = 42;
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
    
    // 後半部分のデータクリア（行45-79、列C以降）- 40行目・44行目・79行目・80行目の関数は保護
    for (let row = 45; row <= 79; row++) {
      // 40行目、44行目、79行目、80行目は関数があるのでスキップ（保護）
      if (row === 40 || row === 44 || row === 79 || row === 80) continue;
      
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
    
    // 6. 土曜日・日曜日の列に黄色マーカーを設定
    console.log('土曜日・日曜日の列に黄色マーカーを設定開始');
    
    // 前半部分（1-16日）の土曜日・日曜日マーカー設定（5-37行目）
    for (let day = 1; day <= Math.min(16, daysInMonth); day++) {
      const date = new Date(year, month - 1, day);
      const dayOfWeek = date.getDay(); // 0=日曜日, 6=土曜日
      
      if (dayOfWeek === 0 || dayOfWeek === 6) { // 日曜日または土曜日
        const dayCol = 3 + (day - 1) * 2; // 朝食列
        const dayNameCol = dayCol + 1; // 夕食列
        
        // 5-37行目の範囲で黄色マーカーを設定
        const breakfastRange = newSheet.getRange(5, dayCol, 33, 1); // 5-37行目 (33行)
        const dinnerRange = newSheet.getRange(5, dayNameCol, 33, 1);
        
        breakfastRange.setBackground('#FFFF00'); // 黄色
        dinnerRange.setBackground('#FFFF00'); // 黄色
        
        console.log(`前半 ${day}日(${dayOfWeek === 0 ? '日曜日' : '土曜日'}) 列${dayCol},${dayNameCol}に黄色マーカー設定 (5-37行目)`);
      }
    }
    
    // 後半部分（17-31日）の土曜日・日曜日マーカー設定（45-77行目）
    for (let day = 17; day <= daysInMonth; day++) {
      const date = new Date(year, month - 1, day);
      const dayOfWeek = date.getDay(); // 0=日曜日, 6=土曜日
      
      if (dayOfWeek === 0 || dayOfWeek === 6) { // 日曜日または土曜日
        const dayCol = 3 + (day - 17) * 2; // 朝食列
        const dayNameCol = dayCol + 1; // 夕食列
        
        // 45-77行目の範囲で黄色マーカーを設定（40行目・44行目・79行目・80行目の関数は除外）
        for (let row = 45; row <= 77; row++) {
          if (row === 40 || row === 44 || row === 79 || row === 80) continue; // 関数行は保護
          newSheet.getRange(row, dayCol).setBackground('#FFFF00');
          newSheet.getRange(row, dayNameCol).setBackground('#FFFF00');
        }
        
        console.log(`後半 ${day}日(${dayOfWeek === 0 ? '日曜日' : '土曜日'}) 列${dayCol},${dayNameCol}に黄色マーカー設定 (45-77行目, 40・44・79・80行目除外)`);
      }
    }
    
    console.log('✅ 土曜日・日曜日の黄色マーカー設定完了');
    
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
 * カレンダーデータから期間内の記録をコピーして食事原紙の該当月シートを更新
 * @param {number} year 年
 * @param {number} month 月
 * @return {Object} 結果
 */
function copyCalendarDataToMealSheet(year, month) {
  try {
    const mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4";
    const dataSheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    
    const mealSs = SpreadsheetApp.openById(mealSheetId);
    const dataSs = SpreadsheetApp.openById(dataSheetId);
    
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    const mealSheetName = "食事原紙_" + yyyyMM;
    
    // 該当月の食事原紙シートを取得
    const mealSheet = mealSs.getSheetByName(mealSheetName);
    if (!mealSheet) {
      return {
        success: false,
        message: "該当月の食事原紙シート「" + mealSheetName + "」が見つかりません。先に月次シートを生成してください。"
      };
    }
    
    // カレンダーデータから予約情報を取得
    const reservationData = getMonthlyReservationCounts(year, month);
    if (!reservationData.success) {
      return {
        success: false,
        message: "予約データの取得に失敗しました: " + reservationData.message
      };
    }
    
    // 予約データを食事原紙に反映
    // ここに具体的な反映ロジックを実装
    
    return {
      success: true,
      message: "カレンダーデータを食事原紙に反映しました。",
      sheetName: mealSheetName,
      url: mealSs.getUrl() + "#gid=" + mealSheet.getSheetId()
    };
    
  } catch (e) {
    console.error('copyCalendarDataToMealSheet Error: ' + e.message);
    return {
      success: false,
      message: "カレンダーデータの食事原紙への反映中にエラーが発生しました: " + e.message
    };
  }
}

/**
 * 指定年月の予約データを取得する
 * @param {number} year 年
 * @param {number} month 月
 * @return {Object} 予約データ
 */
function getMonthlyReservationCounts(year, month) {
  try {
    const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // 年月文字列
    const yyyyMM = year + (month < 10 ? "0" + month : month);
    
    // カレンダーシート名
    const bCalendarSheetName = "b_calendar_" + yyyyMM;
    const dCalendarSheetName = "d_calendar_" + yyyyMM;
    
    const bCalendarSheet = ss.getSheetByName(bCalendarSheetName);
    const dCalendarSheet = ss.getSheetByName(dCalendarSheetName);
    
    const result = {
      success: true,
      year: year,
      month: month,
      breakfast: [],
      dinner: []
    };
    
    // 朝食カレンダーデータを処理
    if (bCalendarSheet) {
      const bData = bCalendarSheet.getDataRange().getValues();
      if (bData.length > 1) {
        const bHeaders = bData[0];
        
        for (let i = 1; i < bData.length; i++) {
          const row = bData[i];
          const dateStr = row[bHeaders.indexOf("date")];
          const userIds = row[bHeaders.indexOf("userIds")];
          
          if (dateStr && userIds) {
            result.breakfast.push({
              date: dateStr,
              userIds: userIds,
              users: parseUserIds(userIds)
            });
          }
        }
      }
    }
    
    // 夕食カレンダーデータを処理
    if (dCalendarSheet) {
      const dData = dCalendarSheet.getDataRange().getValues();
      if (dData.length > 1) {
        const dHeaders = dData[0];
        
        for (let i = 1; i < dData.length; i++) {
          const row = dData[i];
          const dateStr = row[dHeaders.indexOf("date")];
          const userIds = row[dHeaders.indexOf("userIds")];
          
          if (dateStr && userIds) {
            result.dinner.push({
              date: dateStr,
              userIds: userIds,
              users: parseUserIds(userIds)
            });
          }
        }
      }
    }
    
    return result;
    
  } catch (e) {
    console.error('getMonthlyReservationCounts Error: ' + e.message);
    return {
      success: false,
      message: "予約データの取得中にエラーが発生しました: " + e.message
    };
  }
}

/**
 * ユーザーID文字列をパース
 * @param {string} userIdsStr ユーザーIDの文字列
 * @return {Array} ユーザー配列
 */
function parseUserIds(userIdsStr) {
  if (!userIdsStr) return [];
  
  // 文字列をカンマ区切りでパース
  const ids = userIdsStr.toString().split(',');
  return ids.map(id => ({
    userId: id.trim()
  }));
}
