function transposeArray(inputArray) {
  var numRows = inputArray.length;
  var numCols = inputArray[0].length;

  // 新しい2次元配列を作成し、行と列を入れ替える
  var transposedArray = [];
  for (var i = 0; i < numCols; i++) {
    transposedArray[i] = [];
    for (var j = 0; j < numRows; j++) {
      transposedArray[i][j] = inputArray[j][i];
    }
  }

  return transposedArray;
}


/**
 * 月で最初の月曜日の日付を取得する関数
 * @return {Date} firstMonday
 */
function calcFirstMonday() {
  const date = new Date(); //今日の日付
  const day = 1; //月曜日

  const year = date.getFullYear();
  const month = date.getMonth();

  for (let i = 1; i <= 7; i++){
    const tmpDate = new Date(year, month, i);

    if (month !== tmpDate.getMonth()) break; //月代わりで処理終了
    if (tmpDate.getDay() !== day) continue; //引数に指定した曜日以外の時は何もしない
    const firstMonday = tmpDate;
    return firstMonday;
  }
}


/**
 * 今日が第何週か求める関数
 * 月曜始まり
 */
function getWeekNumber(){
  const today = new Date();
  const firstMonday = calcFirstMonday();
  let weekNumer;
  const diffDays = today.getDate() - firstMonday.getDate();
  if(diffDays < 0){
    weekNumer = 1;
  }else{
    weekNumer = Math.floor((diffDays/7)) + 2;
  }

  console.log(weekNumer);

  return weekNumer;
}