/**
 * 今日がn月のn週目と表示する関数
 * @param {Date} today
 * @return {Array} [weekNumber,mondayOfWeek]
 */
function updateWeekNumber(today){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName('設定');

  const weekNumberInfo = getWeekNumber(today);
  const year = weekNumberInfo[0];
  const month = weekNumberInfo[1]+1;
  const weekNumber = weekNumberInfo[2];
  const mondayOfWeek = weekNumberInfo[3];

  sheet.getRange('B2').setValue(`${Utilities.formatDate(mondayOfWeek,'JST','yyyy/MM/dd')}時点`);
  sheet.getRange('B14').setValue(year);
  sheet.getRange('B15').setValue(month);

  // console.log(`${year.toString().substring(2)}年${month}月${weekNumber}週目`);
  // console.log(weekNumber);
  return [weekNumber,mondayOfWeek];
}

/**
 * 最終更新日時を記録する関数
 */
function recordLatestUpdate(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = spreadSheet.getSheetByName('設定');
  settingSheet.getRange('B3').setValue(new Date());
}


/**
 * 今日が第何週か求める関数
 * 月曜始まり
 * n月最初の月曜の週を第一週とする
 * @param {Array} [yyyy年,mm月,n週目,その週の月曜日 ]
 */
function getWeekNumber(date){
  const mondayOfWeek = getMondayOfWeek(date);
  const firstMonday = calcFirstMonday(mondayOfWeek.getFullYear(),mondayOfWeek.getMonth());
  const diff = mondayOfWeek.getDate() - firstMonday.getDate();
  const weekNumber = diff/7 + 1;

  return [firstMonday.getFullYear(),firstMonday.getMonth(),weekNumber,mondayOfWeek];
}

/**
 * その週の月曜日を取得する
 * @param {Date} date
 */
function getMondayOfWeek(date) {

  // 月曜日を取得する
  let n = 0;

  // 本日
  const today = date; // 月曜日を設定

  // 月曜日を取得
  const mon = today.getDate() - (today.getDay() === 0 ? 6 : today.getDay() - 1);

  // 日数を加算
  const x = mon + n;

  // 日を設定
  const dayOfWeek = new Date(today.setDate(x));

  return dayOfWeek;
}

/**
 * 月で最初の月曜日の日付を取得する関数
 * @param {Number} year
 * @param {Number} month
 * @return {Date} firstMonday
 */
function calcFirstMonday(year,month) {
  const day = 1; //月曜日

  for (let i = 1; i <= 7; i++){
    const tmpDate = new Date(year, month, i);

    if (month !== tmpDate.getMonth()) break; //月代わりで処理終了
    if (tmpDate.getDay() !== day) continue; //引数に指定した曜日以外の時は何もしない
    const firstMonday = tmpDate;

    return firstMonday;
  }
}

function checkWKFunction(){
  const today = new Date("2024/02/04");
  const weekNum = getWeekNumber(today);
  console.log(weekNum);
}