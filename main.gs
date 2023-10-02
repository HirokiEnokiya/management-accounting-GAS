
/**
 * 起動時に実行する関数
 */
function onOpenFunction(){
  const today = new Date("2023/10/02");
  adjustSheetsToWeekNumber(today);
  // 当月見込
  updateThisMonthProspectData(today,false,"prospect");
  // 当月予算
  updateThisMonthProspectData(today,false,"budget");
  // 来月見込
  updateThisMonthProspectData(today,true,"prospect");
  // 来月予算
  updateThisMonthProspectData(today,true,"budget");

  // updateThisMonthProspectData(today);
  // updateNextMonthProspectData(today);
  recordLatestUpdate();

}

/**
 * 週に対応してシートの中身を変える関数
 * @pram {Date} today
 */
function adjustSheetsToWeekNumber(today){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadSheet.getSheets();
  const weekNumber = updateWeekNumber(today)[0];

  // 設定シートの更新
  setOutputColumnNumber(weekNumber);

  for(sheet of sheets){
    const sheetName = sheet.getName();
    if(sheetName !== 'template' && sheetName !== '設定' && sheetName !== 'メンテナンス'){
      if(weekNumber === 1){
        sheet.getRange(4,3,sheet.getLastRow()-3,sheet.getLastColumn()-2).clearContent();
        // 0で埋める
        const zeroArray = new Array(143).fill([0,0,0,0]);
        sheet.getRange(4,3,sheet.getLastRow()-3,4).setValues(zeroArray);
        // 数式をセット
        sheet.getRange('G1').setFormula(`='設定'!B9`);
        sheet.getRange('M1').setFormula(`='設定'!B10`);
        sheet.getRange('S1').setFormula(`='設定'!B11`);
        sheet.getRange('Y1').setFormula(`='設定'!B12`);
        sheet.getRange('AE1').setFormula(`='設定'!B13`);
        SpreadsheetApp.flush();
      }else if(weekNumber === 2){
        sheet.getRange(4,3,sheet.getLastRow()-3,4).clearContent();
        // 0で埋める
        const zeroArray = new Array(143).fill([0,0]);
        sheet.getRange(4,5,sheet.getLastRow()-3,2).setValues(zeroArray);
        // 数式をセット
        sheet.getRange('C4').setFormula(`=ARRAYFORMULA(INDIRECT('設定'!$B$5&4&":"&'設定'!$B$5&146) - G4:G146)`);
        sheet.getRange('D4').setFormula(`=ARRAYFORMULA(INDIRECT('設定'!$C$5&4&":"&'設定'!$C$5&146) - J4:J146)`);
      }
      else{
        // 数式をセット
        sheet.getRange(4,3,sheet.getLastRow()-3,4).clearContent();
        sheet.getRange('C4').setFormula(`=ARRAYFORMULA(INDIRECT('設定'!$B$5&4&":"&'設定'!$B$5&146) - G4:G146)`);
        sheet.getRange('D4').setFormula(`=ARRAYFORMULA(INDIRECT('設定'!$C$5&4&":"&'設定'!$C$5&146) - J4:J146)`);
        sheet.getRange('E4').setFormula(`=ARRAYFORMULA(INDIRECT('設定'!$B$5&4&":"&'設定'!$B$5&146) - INDIRECT('設定'!$B$7&4&":"&'設定'!$B$7&146))`);
        sheet.getRange('F4').setFormula(`=ARRAYFORMULA(INDIRECT('設定'!$C$5&4&":"&'設定'!$C$5&146) - INDIRECT('設定'!$C$7&4&":"&'設定'!$C$7&146))`);

      }
    }
  }
}

/**
 * 見込進捗確認シートでその週に対応する列を設定シートにセットする関数
 * @param {Number} weekNumber
 */
function setOutputColumnNumber(weekNumber){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = spreadSheet.getSheetByName('設定');
  const thisMonthTargetColumnNum = weekNumber*6 + 1;
  // 参照する列の変更
  settingSheet.getRange('B4').setValue(thisMonthTargetColumnNum);
}



/**
 * 今月or来月の見込or予算について、予実管理シートの値をこのスプレッドシートに反映する関数
 * @param {Date} today
 * @param {Boolean} isNextMonth true(来月) flase(今月)
 * @param {String} category "prospect" or "budget"
 */
function updateThisMonthProspectData(today,isNextMonth,category){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = spreadSheet.getSheetByName('設定');
  const officeName = settingSheet.getRange('B1').getValue();

  // その日が第何週か計算
  const weekNumberInfo = updateWeekNumber(today);
  const weekNumber = weekNumberInfo[0];
  const mondayOfWeek = weekNumberInfo[1];

  // この事業部の予実管理シートを取得
  const budgetControlSpreadSheetId = getBudgetControlSpreadSheetId();
  const spreadSheetId = budgetControlSpreadSheetId[officeName];
  const budgetControlSheet = SpreadsheetApp.openById(spreadSheetId);

  // 月とカテゴリによる列番号の差分を計算
  let firstInputColumnNumber; //予実管理シートの最初の月での抜き出す項目の列番号
  let relativeOutputColumnIndex; //値をセットする列がその週のフィールドの中で何番目にあるか(0始まり)
  switch(category){
    case "prospect":
      firstInputColumnNumber = 3;
      relativeOutputColumnIndex = 0 + Number(isNextMonth)*3;
      break;
    case "budget":
      firstInputColumnNumber = 5;
      relativeOutputColumnIndex = 1 + Number(isNextMonth)*3;
      break;
    default:
      throw new Error("Invalid category");
  }

  // 予実管理シートの抽出する列を計算
  let month = mondayOfWeek.getMonth() + 1;
  if(month === 1||month === 2||month === 3){
    month += 12;
  }
  const thisMonth = month + Number(isNextMonth);
  const inputColumnNumber = firstInputColumnNumber + (thisMonth - 4)*7;

  // このシートの反映する列を週ごとに変える
  const outputColumnNumber = weekNumber*6 + 1 + relativeOutputColumnIndex;

  // 値の取得と反映
  for(sheet of spreadSheet.getSheets()){
    if(sheet.getName() !== 'template' && sheet.getName() !==  "設定" && sheet.getName() !== "メンテナンス"){
      const shipper = sheet.getName();
      try{
        const sourceSheet = budgetControlSheet.getSheetByName(shipper);
        const columnData = sourceSheet.getRange(4,inputColumnNumber,sourceSheet.getLastRow()-3,1).getDisplayValues();

        sheet.getRange(4,outputColumnNumber,columnData.length,1).setValues(columnData);
      }catch(e){
        console.log(e);
        console.log(`${shipper}のシートが見つかりません`);
      }
    }

  }
}

