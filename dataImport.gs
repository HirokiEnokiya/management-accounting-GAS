/**
 * 荷主マスタのデータを管理シートから取得する関数
 * @return {Array} 荷主マスタのデータのうちその事業所配下のもの
 */
function getOfficeShipperList(){
  let shipperList = {};
  const officeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const officeName = officeSpreadSheet.getSheetByName('設定').getRange('B1').getValue();
  console.log(officeName);
  const controlSpreadSheetId = officeSpreadSheet.getSheetByName('設定').getRange('B16').getValue();
  const controlSpreadSheet = SpreadsheetApp.openById(controlSpreadSheetId);
  const shipperMasterSheet = controlSpreadSheet.getSheetByName('荷主マスタ');
  if(shipperMasterSheet.getLastRow() < 1 || shipperMasterSheet.getLastColumn() < 1){
    throw new Error("荷主マスタに情報が登録されていません");
  }
  let table = shipperMasterSheet.getRange(2,2,shipperMasterSheet.getLastRow()-1,shipperMasterSheet.getLastColumn()-1).getValues();
  table = transposeArray(table);
  for(let i=0;i<table.length;i++){
    let officeData = table[i];
    const officeName = officeData.shift();
    shipperList[officeName] = officeData.filter(Boolean); //空要素を削除
  }
  return shipperList[officeName];
}

/**
 * シート上の事業部名に対応する荷主データを外部シートから取ってくる関数
 * @param {String} officeName
 * @return {Object} shippersData
 */
function importData(officeName){
  const SOURCE_SS_ID = PropertiesService.getScriptProperties().getProperty('SOURCE_SS_ID');
  const sourceDataSpreadSheet = SpreadsheetApp.openById(SOURCE_SS_ID);
  const sourceSheet = sourceDataSpreadSheet.getSheetByName(officeName);
  
  // 取得したデータをオブジェクトにする
  let tableData = sourceSheet.getRange(1,3,sourceSheet.getLastRow(),sourceSheet.getLastColumn()-2).getValues();
  tableData = transposeArray(tableData);

  let shippersData = {};

  for(let i=0;i<tableData.length;i++){
    let shipperData = tableData[i];
    const shipperName = shipperData[0];
    shipperData.shift();
    shippersData[shipperName] =  shipperData;
  }

  return shippersData;

}