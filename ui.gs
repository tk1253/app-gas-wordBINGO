// スプレッドシートが開かれたら自動的に実行される関数
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('[BINGO]')
    .addItem('ガラガラポン!', 'DoGaragarapon')
    .addSeparator()
    .addSubMenu(
      ui.createMenu("設定")
        .addItem('ビンゴカード配布', 'createBingoCards')
        .addItem('ガラポン初期化', 'reset')
        .addItem('ビンゴカード削除', 'deleteBingoCards'))
    .addToUi();
}
