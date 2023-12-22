function deleteBingoCards() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();

  //シート名に「BingoCard」が含まれるシートを削除する
  for (let i = 0; i < sheets.length; i++) {
    let sheetName = sheets[i].getName();
    if (sheetName.indexOf('BingoCard') !== -1) {
      ss.deleteSheet(sheets[i]);
      Logger.log('BingoCard sheet deleted: ' + sheetName);
    }
  }
}
