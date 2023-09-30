/**
 * 今日がn月のn週目と表示する関数
 * @return {Number} weekNumber
 */
function updateWeekNumber(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName('設定');

  const today = new Date("2023/10/09");
  // デバッグ用
  // const today = new Date();
  let year = today.getFullYear();
  const month = today.getMonth() + 1;
  const weekNumber = getWeekNumber(today);

  sheet.getRange('B2').setValue(`'${year.toString().substring(2)}年${month}月${weekNumber}週目`);
  sheet.getRange('B14').setValue(year);
  sheet.getRange('B15').setValue(month);
  return weekNumber;
}

/**
 * 最終更新日時を記録する関数
 */
function recordLatestUpdate(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = spreadSheet.getSheetByName('設定');
  settingSheet.getRange('B3').setValue(new Date());
}