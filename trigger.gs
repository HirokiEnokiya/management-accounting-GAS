/**
 * 初期設定
 */
function initialize(){
  //起動時に実行するトリガーを作成
  ScriptApp.newTrigger("onOpenFunction")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()  //スプレッドシートを開いた時
      .create();
  Browser.msgBox('初期設定が完了しました。再読み込みをしてください。');
  
}