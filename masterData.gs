/**
 * 事業所名称の一覧をシートから取得する関数
 * 荷主や事業所が増えたらスプシを編集する
 * @return {Object} shipperList
 */
function getShipperList(){
  let shipperList = {};
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName('荷主一覧');
  let table = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  table = transposeArray(table);
  for(let i=0;i<table.length;i++){
    let officeData = table[i];
    const officeName = officeData.shift();
    shipperList[officeName] = officeData.filter(Boolean); //空要素を削除
  }
  return shipperList;

}

/**
 * 各事業所のスプレッドシートのIDをシートから取得する関数
 * @return {Object} officeSpreadSheetIds
 */
function getOfficeSpreadSheetIds(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName('スプレッドシートID');
  const table = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
  const officeSpreadSheetIds = Object.fromEntries(table);
  return officeSpreadSheetIds;
}