function createBingoCards() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("要素");

  // シートが存在しなければログを吐いて終了
  if (!sheet) {
    Logger.log('Sheet "要素" not found.');
    return;
  }

  // シートから要素（テキスト）のリストを取得
  let names = getNamesFromSheet(sheet);

 // 名前が取得できなければログを吐いて終了
  if (names.length === 0) {
    Logger.log('No names found in sheet "要素".');
    return;
  }

  let numRows = 5;
  let numColumns = 5;

 // 新規シートとしてビンゴカードを8枚作成する
 // 要素のリストをシャッフルする
  for (let cardNumber = 1; cardNumber <= 8; cardNumber++) {
    let shuffledNames = fisherYatesShuffle([...names]);

    let newSheet = ss.insertSheet('BingoCard_' + cardNumber);

    if (!newSheet) {
      Logger.log('Error creating sheet for BingoCard_' + cardNumber);
      continue;
    }

    newSheet.clear();

    // ビンゴカードのセル（5×5）に、シャッフルされた要素を配置するループ
    for (let i = 0; i < numRows; i++) {
      for (let j = 0; j < numColumns; j++) {
        let cell = newSheet.getRange(i + 1, j + 1);
        cell.setValue(shuffledNames[i * numColumns + j]);
        cell.setHorizontalAlignment("center");
        cell.setVerticalAlignment("middle");

        if (cell.getValue() === "FREE") {
          cell.setBackgroundColor("#F4CCCC");
        }
      }
    }

    // 中央のセルをFREEに設定し、背景色と行列幅を変更
    let centerCell = newSheet.getRange(Math.floor(numRows / 2) + 1, Math.floor(numColumns / 2) + 1);
    centerCell.setValue("FREE");
    centerCell.setHorizontalAlignment("center");
    centerCell.setVerticalAlignment("middle");
    centerCell.setBackgroundColor("#F4CCCC");

    newSheet.setRowHeights(1, numRows, 100);
    newSheet.setColumnWidths(1, numColumns, 200);
  }
}

// Fisher-Yates シャッフルのアルゴリズム関数
function fisherYatesShuffle(array) {
  let currentIndex = array.length, temporaryValue, randomIndex;

  while (0 !== currentIndex) {
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex -= 1;

    temporaryValue = array[currentIndex];
    array[currentIndex] = array[randomIndex];
    array[randomIndex] = temporaryValue;
  }

  return array;
}

// シートから要素（テキスト）のリストを取得する関数
function getNamesFromSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("要素");
  
  // Assuming names are in column A, starting from the second row
  let range = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
  let values = range.getValues();
  
  // Flatten the 2D array into a 1D array
  let names = values.flat();
  
  return names;
}

