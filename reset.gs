function reset() {
  
  // 済マークを消す
  let num_item_sheet = SpreadsheetApp.getActive().getSheetByName('要素');
  let num_item_data = num_item_sheet.getDataRange().getValues();
  for(i = 1; i < num_item_data.length; i++) {
    num_item_sheet.getRange(i + 1, 3).setValue('');
  }

  // ビンゴを消す
  let bingo_sheet = SpreadsheetApp.getActive().getSheetByName('ビンゴ');
  let bingo_data = num_item_sheet.getDataRange().getValues();
  for(i = 0; i < bingo_data.length; i++) {
    bingo_sheet.getRange(i + 1, 1).setValue('');
    bingo_sheet.getRange(i + 1, 2).setValue('');
  }

}
