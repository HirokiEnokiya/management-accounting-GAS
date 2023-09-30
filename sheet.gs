/**
 * 荷主の増減に従ってシートを追加・削除する関数
 * @return {Object} changedSheetNames
 */
function cleanSheets() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = spreadSheet.getSheetByName('template');
  const setSheet = spreadSheet.getSheetByName('設定');
  const officeName = setSheet.getRange('B1').getValue();

  // 現在のシート名
  const sheets = spreadSheet.getSheets();
  let sheetNames = sheets.map(sheet => sheet.getName());
  sheetNames = sheetNames.filter(function(name){
    return (name !== '設定' && name !== 'template' && name !== 'メンテナンス')
  })

  // 最新の荷主マスタのデータを取得
  const shipperList = getOfficeShipperList();

  // 荷主が増えていないか
  const addedShippers = shipperList.filter(function(shipper){
    return !sheetNames.includes(shipper);
  });
  console.log(addedShippers);
  // 増えていたらシートを追加
  if(addedShippers.length > 0){
    for(shipperName of addedShippers){
        const copiedSheet = templateSheet.copyTo(spreadSheet);
        copiedSheet.setName(shipperName);
        copiedSheet.getRange('A1').setValue(`${officeName}${shipperName}`);
    }
  }

  // 荷主が減っていないか
  const deletedShippers = sheetNames.filter(function(sheetName){
    return !shipperList.includes(sheetName);
  });
  console.log(deletedShippers);
  // 減っていたらシートを削除
  if(deletedShippers.length > 0){
    for(shipperName of deletedShippers){
        const targetSheet = spreadSheet.getSheetByName(shipperName);
        spreadSheet.deleteSheet(targetSheet);
    }
  }
  
  const changedSheetNames ={
    'added':addedShippers,
    'deleted':deletedShippers,
  }

  return changedSheetNames;
}




function checkFunc(){
  const list = getOfficeShipperList();
  console.log(list);
}