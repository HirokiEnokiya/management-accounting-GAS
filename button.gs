function updateShipperSheets() {
  let response = Browser.msgBox("追加された荷主のシートの作成と、削除された荷主のシートの削除を行います。",Browser.Buttons.OK_CANCEL);
  if(response == "ok"){
    const changedSheetNames = cleanSheets();

    Browser.msgBox(`${changedSheetNames.added}シートを追加しました。\\n${changedSheetNames.deleted}シートを削除しました。`);

  }

}
