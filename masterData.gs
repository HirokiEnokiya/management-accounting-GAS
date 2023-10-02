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