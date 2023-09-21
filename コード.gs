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


/**
 * 各事業所ごとのシートを作る関数
 * @param {Object} shipperList
 */
function makeSheet(shipperList){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const tableTemplateSheet = spreadSheet.getSheetByName('見出し');

  const officeNameList = Object.keys(shipperList);

  for(let i=0;i<officeNameList.length;i++){
    const copiedSheet = tableTemplateSheet.copyTo(spreadSheet);

    const officeName = officeNameList[i];
    copiedSheet.setName(officeName);
    // const shippers = shipperList[officeName];
    // copiedSheet.getRange(1,3,1,shippers.length).setValues([shippers]);
  }

}

/**
 * いらなくなったシートを削除する関数
 * @param {Array} sheetNames
 */
function deleteSheets(sheetNames){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  for(sheetName of sheetNames){
    const sheet = spreadSheet.getSheetByName(sheetName);
    spreadSheet.deleteSheet(sheet);
  }
}

/**
 * 月に対応するスプレッドシートの列を計算する関数
 * @return {Number} columnNum
 */
function getColumnNum(){
  const today = new Date();
  let month = today.getMonth();

  if(month === 1||2||3){
    month += 12;
  }

  const columnNum = 3 + (month - 4)*7;
  return columnNum;
}

/**
 * n月第n週を表記する関数
 */
function showWK(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName('設定');

  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const weekNumber = getWeekNumber();

  sheet.getRange('A1').setValue(`${year}年 ${month}月 第${weekNumber}週`);
}

// データ取得系
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
 * テスト
 */
function funcCheck(){
  const spreadSheetId = '1rkXM0ni7Go1keGWOpErv5FAjnvw2nyl3Q4RJEZ787fQ';
  const columnNum = getColumnNum();

  console.log(mainFunction());
}

// function tempFunc(){
//   let column = 6;
//   const officeSpreadSheetIds = getOfficeSpreadSheetIds();
//   for(officeName in officeSpreadSheetIds){
//     const officeSpreadSheetId = officeSpreadSheetIds[officeName];
//     const spreadSheet = SpreadsheetApp.openById(officeSpreadSheetId);
//     let array = [];
//     const sheets = spreadSheet.getSheets();
//     for(i=0;i<sheets.length;i++){
//       const sheetName = sheets[i].getName();
//       array.push([sheetName]);
//     }
//     SpreadsheetApp.getActiveSpreadsheet().getSheetByName('荷主一覧').getRange(2,column,array.length,1).setValues(array);
//     column++;
//   }
// }