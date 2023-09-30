/**
 * 初期設定
 */
function initialize(){
  //起動時に実行するトリガーを作成
  ScriptApp.newTrigger("onOpenFunction")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()  //スプレッドシートを開いた時
      .create();
  
}

/**
 * 月が変わったらシートの中身を全部リセットする関数
 */
function resetSheets(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadSheet.getSheets();
  for(sheet of sheets){
    const sheetName = sheet.getName();
    if(sheetName !== 'template' && sheetName !== '設定' && sheetName !== 'メンテナンス'){
      sheet.getRange(4,3,sheet.getLastRow()-3,sheet.getLastColumn()-2).clearContent();
      // 数式をセット
      sheet.getRange('C4').setFormula(`=ARRAYFORMULA(INDIRECT('設定'!$B$5&4&":"&'設定'!$B$5&146) - G4:G146)`);
      sheet.getRange('E4').setFormula(`=ARRAYFORMULA(INDIRECT('設定'!$B$5&4&":"&'設定'!$B$5&146) - INDIRECT('設定'!$B$7&4&":"&'設定'!$B$7&146))`);
      sheet.getRange('M4').setFormula(`=ARRAYFORMULA(G4:G146)`);
      sheet.getRange('S4').setFormula(`=ARRAYFORMULA(M4:M146)`);
      sheet.getRange('Y4').setFormula(`=ARRAYFORMULA(S4:S146)`);
      sheet.getRange('AE4').setFormula(`=ARRAYFORMULA(Y4:Y146)`);
      sheet.getRange('AK4').setFormula(`=ARRAYFORMULA(AE4:AE146)`);
      sheet.getRange('G1').setFormula(`='設定'!B9`);
      sheet.getRange('M1').setFormula(`='設定'!B10`);
      sheet.getRange('S1').setFormula(`='設定'!B11`);
      sheet.getRange('Y1').setFormula(`='設定'!B12`);
      sheet.getRange('AE1').setFormula(`='設定'!B13`);
      sheet.getRange('AK1').setFormula(`='設定'!B14`);
    }
  }
}

/**
 * 起動時に実行する関数
 */
function onOpenFunction(){
  const weekNumber = updateWeekNumber();
  if(weekNumber === 1){
    resetSheets();
    SpreadsheetApp.flush();
  }
  updateThisMonthProspectData();
  updateNextMonthProspectData();

}



/**
 * 当月見込を反映する関数
 */
function updateThisMonthProspectData(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = spreadSheet.getSheetByName('設定');
  const weekNumber = updateWeekNumber();
  const nextMonthTargetColumnNum = getThisMonthTargetColumnNum(weekNumber);
  const officeName = settingSheet.getRange('B1').getValue();
  const budgetControlSpreadSheetId = getBudgetControlSpreadSheetId();
  // 予実管理シート
  const spreadSheetId = budgetControlSpreadSheetId[officeName];
  const budgetControlSheet = SpreadsheetApp.openById(spreadSheetId);

  const today = new Date();
  let month = today.getMonth() + 1;
  if(month === 1||month === 2||month === 3){
    month += 12;
  }
  const thisMonth = month;
  const sourceTargetColumnNum = 3 + (thisMonth - 4)*7;

  for(sheet of spreadSheet.getSheets()){
    if(sheet.getName() !== 'template' && sheet.getName() !==  "設定" && sheet.getName() !== "メンテナンス"){
      const shipper = sheet.getName();
      try{
        const sourceSheet = budgetControlSheet.getSheetByName(shipper);
        const columnData = sourceSheet.getRange(4,sourceTargetColumnNum,sourceSheet.getLastRow()-3,1).getDisplayValues();

        sheet.getRange(4,nextMonthTargetColumnNum,columnData.length,1).setValues(columnData);
      }catch(e){
        console.log(e);
        console.log(`${shipper}のシートが見つかりません`);
      }
    }

  }
}

/**
 * 翌月見込を反映する関数
 */
function updateNextMonthProspectData(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = spreadSheet.getSheetByName('設定');
  const weekNumber = updateWeekNumber();
  const nextMonthTargetColumnNum = getThisMonthTargetColumnNum(weekNumber) + 3;
  const officeName = settingSheet.getRange('B1').getValue();
  const budgetControlSpreadSheetId = getBudgetControlSpreadSheetId();
  // 予実管理シート
  const spreadSheetId = budgetControlSpreadSheetId[officeName];
  const budgetControlSheet = SpreadsheetApp.openById(spreadSheetId);

  const today = new Date();
  let month = today.getMonth() + 1;
  if(month === 1||month === 2||month === 3){
    month += 12;
  }
  const nextMonth = month + 1;
  const sourceTargetColumnNum = 3 + (nextMonth - 4)*7;

  for(sheet of spreadSheet.getSheets()){
    if(sheet.getName() !== 'template' && sheet.getName() !==  "設定" && sheet.getName() !== "メンテナンス"){
      const shipper = sheet.getName();
      try{
        const sourceSheet = budgetControlSheet.getSheetByName(shipper);
        const columnData = sourceSheet.getRange(4,sourceTargetColumnNum,sourceSheet.getLastRow()-3,1).getDisplayValues();

        sheet.getRange(4,nextMonthTargetColumnNum,columnData.length,1).setValues(columnData);
      }catch(e){
        console.log(e);
        console.log(`${shipper}のシートが見つかりません`);
      }
    }

  }
}

/**
 * 事業所ごとの予実管理シートのidを取得する関数
 * @return {Object} budgetControlSpreadSheetIds
 */
function getBudgetControlSpreadSheetId(){
  let budgetControlSpreadSheetIds = {};
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = spreadSheet.getSheetByName('設定');
  const table = settingSheet.getRange(5,6,settingSheet.getLastRow(),2).getValues();
  for(row of table){
    const officeName = row[0];
    const id = row[1];
    if(officeName !== ""){
      budgetControlSpreadSheetIds[officeName] = id;
    }
  }
  return budgetControlSpreadSheetIds;
  }


// /**
//  * データ一時保管シートから反映する関数
//  */
// function updateThisMonthData(){
//   const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
//   const settingSheet = spreadSheet.getSheetByName('設定');
//   const weekNumber = updateWeekNumber();
//   const officeName = settingSheet.getRange('B1').getValue();
//   const targetColumnNum = getThisMonthTargetColumnNum(weekNumber);

//   // 月初めならばシートをリセットする
//   if(weekNumber === 1){
//     resetSheets();
//   }

//   const shippersData = importData(officeName);

//   for(shipperName in shippersData){
//     try{
//       const targetSheet = spreadSheet.getSheetByName(shipperName);
//       const outputData = shippersData[shipperName];
//       targetSheet.getRange(4,targetColumnNum,outputData.length,1).setValues(transposeArray([outputData]));
//     }catch(e){
//       console.log(e);
//       console.log(`${shipperName}のシートがみつかりません`);
//     }

//   }

//   // 参照する列の変更
//   settingSheet.getRange('B4').setValue(targetColumnNum);

// }





/**
 * 月に対応するスプレッドシートの列を計算する関数
 * @param {Number} weekNumber
 * @return {Number} columnNum
 */
function getThisMonthTargetColumnNum(weekNumber){
  const columnNum = weekNumber*6 + 1;
  return columnNum
}
