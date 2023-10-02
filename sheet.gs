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
    try{
      const sheet = spreadSheet.getSheetByName(sheetName);
      spreadSheet.deleteSheet(sheet);
    }catch{
      console.log(`${sheetName}シートが存在しません`);
    }
  }
}