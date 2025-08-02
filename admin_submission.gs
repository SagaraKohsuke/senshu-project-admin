/**
 * 月次の予約データを「食事原紙」をテンプレートとして新しいスプレッドシートに出力する
 * (前半・後半でブロックが分かれた特殊なレイアウトに対応)
 * @param {number} year 年
 * @param {number} month 月
 * @return {Object} 成功した場合はスプレッドシートのURLを含むオブジェクト、失敗した場合はエラーメッセージを含むオブジェクト
 */
function exportMonthlyReservationsToSheet(year, month) {
  try {
    // ユーザーシートからIDと名前の対応表を作成
    const usersSheet = ss.getSheetByName("users");
    if (!usersSheet) {
      return { success: false, message: "「users」シートが見つかりません。" };
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
      return { success: false, message: reservationData.message };
    }

    // 「食事原紙」をテンプレートとして新しいシートを作成
    const templateSheet = ss.getSheetByName("食事原紙");
    if (!templateSheet) {
      return { success: false, message: "「食事原紙」という名前のシートが見つかりません。" };
    }
    
    const newSpreadsheetName = `${year}年${month}月 食事申し込み一覧`;
    const newSS = SpreadsheetApp.create(newSpreadsheetName);
    const newSheet = templateSheet.copyTo(newSS);
    newSheet.setName(`${month}月食事`);

    if (newSS.getSheets().length > 1) {
      newSS.deleteSheet(newSS.getSheets()[0]);
    }
    
    // ヘッダー情報（タイトルのみ）を更新
    const titleRange = newSheet.getRange("A1");
    titleRange.setValue(titleRange.getValue().replace('月度', `${month}月度`));

    // 前半・後半ブロックごとにユーザーIDと行のマッピングを作成し、氏名を設定する
    const mapUsersAndSetNames = (startRow, endRow) => {
      const userRowMap = {};
      const idRange = newSheet.getRange(`A${startRow}:A${endRow}`);
      const nameRange = newSheet.getRange(`B${startRow}:B${endRow}`);
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

    // 予約データ（「1」）を正しいブロックのセルに書き込む
    const dataToUpdate = [];
    const { breakfast: breakfastReservations, dinner: dinnerReservations } = reservationData;

    const processReservations = (reservations, isDinner) => {
      reservations.forEach(dayData => {
        if (dayData.users.length > 0) {
          const dayOfMonth = parseInt(dayData.date.split('-')[2], 10);
          
          let userRowMap;
          let relativeDay;

          // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
          //
          //                 ここが修正の核心部分です
          //
          // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
          if (dayOfMonth <= 16) {
            // 前半ブロックの場合
            userRowMap = userRowMap_1_16;
            relativeDay = dayOfMonth;
          } else {
            // 後半ブロックの場合
            userRowMap = userRowMap_17_31;
            // ★列の計算をリセットするために、日付を1から再計算する
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
      newSheet.getRange(data.row, data.col).setValue(data.value);
    });
    
    // 作成したファイルを移動し、URLを返す
    const sourceFile = DriveApp.getFileById(ss.getId());
    const sourceFolder = sourceFile.getParents().next();
    const newFile = DriveApp.getFileById(newSS.getId());
    sourceFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);

    return { success: true, url: newSS.getUrl() };

  } catch (e) {
    console.error('exportMonthlyReservationsToSheet Error: ' + e.message + " Stack: " + e.stack);
    return { success: false, message: "シートの出力中にエラーが発生しました: " + e.message };
  }
}