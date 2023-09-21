/**
 * 各事業所の予実管理シートからその週の責任者予測を抜き出してくる関数
 */
function main(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  // マスタデータの取得
  const shipperList = getShipperList();
  const officeSpreadSheetIds = getOfficeSpreadSheetIds();

  // 過去のシートの削除
  const sheetNames = Object.keys(officeSpreadSheetIds);
  deleteSheets(sheetNames);

  // シートの作成
  makeSheet(shipperList);

  // 各事業所データの反映
  const columnNum = 3;
  for(officeName in officeSpreadSheetIds){

    // 予実管理シートから全シートのデータを取得
    const officeSpreadSheetId = officeSpreadSheetIds[officeName];
    const columnData = getColumnData(officeSpreadSheetId,columnNum);
    console.log(officeName);
    // console.log(columnData);

    // シートに転記
    let shipperValues = [];
    const officeShipperList = shipperList[officeName];
    for(shipper of officeShipperList){
      console.log(shipper);
      try{
        shipperValues.push([shipper,...columnData[shipper]]);
      }catch(e){
        console.log(e);
        console.log(`${shipper}のシートがみつかりません`);
      }
    }
    const outputData = transposeArray(shipperValues);
    const targetSheet = spreadSheet.getSheetByName(officeName);
    targetSheet.getRange(1,3,targetSheet.getLastRow(),shipperValues.length).setValues(outputData);
  }

}

/**
 * 指定したスプレッドシートの全シートの指定した列を抜き出してくる(第4行以降)関数
 * 各事業所の予実管理シートから抜き出す
 * 列番号を与えてその時点での当月の責任者予測の値を荷主ごとに取ってくる
 * @param {String} spreadSheetId
 * @param {Number} columnNum
 * @return {Object} columnsData シート名をキーとして全シートの列データを格納したオブジェクト
 */
function getColumnData(spreadSheetId,columnNum){
  let columnsData = {};
  const spreadSheet = SpreadsheetApp.openById(spreadSheetId);
  for(let i=0;i < spreadSheet.getSheets().length;i++){
    const sheet = spreadSheet.getSheets()[i];
    const sheetName = sheet.getName();
    const columnData = sheet.getRange(4,columnNum,sheet.getLastRow()-3,1).getValues().flat();
    columnsData[sheetName] = columnData;
    // console.log(columnsData);
  }
  return columnsData;
}


