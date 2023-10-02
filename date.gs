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

// /**
//  * n月第n週を表記する関数
//  */
// function showWK(){
//   const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = spreadSheet.getSheetByName('設定');

//   const today = new Date();
//   const year = today.getFullYear();
//   const month = today.getMonth() + 1;
//   const weekNumber = getWeekNumber();

//   sheet.getRange('A1').setValue(`${year}年 ${month}月 第${weekNumber}週`);
// }