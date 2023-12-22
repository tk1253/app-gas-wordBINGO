function DoGaragarapon() {
  // 要素シートから全データを取得
  let all_item_sheet = SpreadsheetApp.getActive().getSheetByName('要素');
  let num_item_data = all_item_sheet.getDataRange().getValues();
  
  // （getSelectedCount関数）すでに選ばれた要素の数を取得
  let selectedCount = getSelectedCount(all_item_sheet);
  
  // 未選択の要素データを取得
  let remainingItems = num_item_data.filter(row => row[2] !== '済');

  // 未選択の要素がない場合はログを出力して終了
  if (remainingItems.length === 0) {
    Logger.log('全ての要素が取得済みです。');
    return;
  }

  // ランダムに要素を1つ選択
  let randomIndex = Math.floor(Math.random() * remainingItems.length);
  let selectedItem = remainingItems[randomIndex];
  let item = selectedItem[1];
  
  // （set_done関数）取得済み要素に済フラグを設定
  set_done(all_item_sheet, item);

  // （show_result関数）結果をビンゴ表示シートに表示する
  let bingo_sheet = SpreadsheetApp.getActive().getSheetByName('ビンゴ');
  show_result(bingo_sheet, selectedCount, item);
}

// 結果をビンゴ表示シートに表示する関数
function show_result(bingo_sheet, selectedCount, item) {
  // 結果を書き出し
  let order = selectedCount + 1;
  bingo_sheet.appendRow([order, item]);
}

// 選ばれた要素に済フラグを入れる関数
function set_done(num_item_sheet, name) {
  let num_item_data = num_item_sheet.getDataRange().getValues();
  for (let i = 1; i < num_item_data.length; i++) {
    if (num_item_data[i][1] == name) {
      num_item_sheet.getRange(i + 1, 3).setValue('済');
      break;
    }
  }
}

// 選ばれた要素の数を取得する関数
function getSelectedCount(num_item_sheet) {
  // シートから全要素データを再取得
  let num_item_data = num_item_sheet.getDataRange().getValues();
  // 済フラグが付いている要素の数をカウント
  let selectedCount = 0;
  for (let i = 1; i < num_item_data.length; i++) {
    if (num_item_data[i][2] == '済') {
      selectedCount++;
    }
  }
  return selectedCount;
}
