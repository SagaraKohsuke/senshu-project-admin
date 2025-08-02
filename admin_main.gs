function doGet() {
  return HtmlService.createTemplateFromFile("admin_index")
    .evaluate()
    .setTitle("泉州会館 管理人画面");
}

const spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk";
const ss = SpreadsheetApp.openById(spreadsheetId);