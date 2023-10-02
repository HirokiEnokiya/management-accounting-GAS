/**
 * 事業所名称の一覧をシートから取得する関数
 * 荷主や事業所が増えたらスプシを編集する
 * @return {Object} shipperList
 */
function getShipperList(){
  let shipperList = {};
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName('荷主マスタ');
  if(sheet.getLastRow() < 1 || sheet.getLastColumn() < 1){
    throw new Error("荷主マスタに情報が登録されていません");
  }
  let table = sheet.getRange(2,2,sheet.getLastRow()-1,sheet.getLastColumn()-1).getValues();
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
  const sheet = spreadSheet.getSheetByName('事業部マスタ');
  const table = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
  const officeSpreadSheetIds = Object.fromEntries(table);
  return officeSpreadSheetIds;
}

/**
 * 荷主マスタを更新する関数
 * 各事業所の予実管理シートのシート名に従う
 */
function updateShipperList(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const shipperListSheet = spreadSheet.getSheetByName('荷主マスタ');
  // 表をリセット
  if(shipperListSheet.getLastRow() > 2 && shipperListSheet.getLastColumn()>1){
    shipperListSheet.getRange(3,2,shipperListSheet.getLastRow()-2,shipperListSheet.getLastColumn()-1).clearContent();
  }
  // 使用しないシート一覧
  const exclusiveSheetNames = shipperListSheet.getRange(2,1,shipperListSheet.getLastRow()-1,1).getValues().flat().filter(Boolean); //空白のセルを削除

  const spreadSheetIds = shipperListSheet.getRange(1,2,1,shipperListSheet.getLastColumn()-1).getValues().flat();
  for(let i=0;i<spreadSheetIds.length;i++){
    const spreadSheetId = spreadSheetIds[i];
    const sourceSpreadSheet = SpreadsheetApp.openById(spreadSheetId);
    const sheets = sourceSpreadSheet.getSheets();
    let sheetNames = sheets.map(sheet => sheet.getName());
    // 除外
    sheetNames = sheetNames.filter(function(name){
      return !exclusiveSheetNames.includes(name);
    });

    console.log(sheetNames);
    shipperListSheet.getRange(3,i+2,sheetNames.length,1).setValues(sheetNames.map(name => [name]));





  }
}