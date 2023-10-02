/**
 * 各事業部ごとの集計スプレッドシートを作成する関数
 */
function makeSpreadSheets(){
  const TEMPLATE_SS_ID = PropertiesService.getScriptProperties().getProperty('TEMPLATE_SS_ID');
  const templateSpreadSheet = SpreadsheetApp.openById(TEMPLATE_SS_ID);
  const FOLDER_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const term = spreadSheet.getSheetByName('設定').getRange('C2').getValue();
  const officeCodes = getOfficeCodes();
  const shipperList = getShipperList();
  console.log(shipperList);

  // 事業部ごとにスプレッドシートを作成
  for(officeName in shipperList){
    const spreadSheetName = `ODK_${term}期_見込進捗確認_${officeCodes[officeName]}_${officeName}`;
    console.log(spreadSheetName);
    const copiedSpreadSheet = templateSpreadSheet.copy(spreadSheetName);
    moveFile(copiedSpreadSheet.getId(),FOLDER_ID);

    // 設定シートの設定
    copiedSpreadSheet.getSheetByName('設定').getRange('B1').setValue(officeName);

    // 荷主ごとにシートを作成
    const shippers = shipperList[officeName];
    for(shipperName of shippers){
      console.log(shipperName);
      const templateSheet = copiedSpreadSheet.getSheetByName('template');
      const copiedSheet = templateSheet.copyTo(copiedSpreadSheet);
      copiedSheet.setName(shipperName);
      copiedSheet.getRange('A1').setValue(`${officeName}${shipperName}`);
    }

  }
}

/**
 * 各事業部のアルファベット対応のオブジェクトを返す関数
 * @return {Object} officeCodes
 */
function getOfficeCodes(){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName('事業部マスタ');
  const table = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
  console.log(table);
  const officeCodes = Object.fromEntries(table.map(function (value){
    return [value[1],value[0]];
  }));

  console.log(officeCodes);
  return officeCodes;
}


/**
 * ファイルを移動させる関数
 * @param {String} fileId
 * @param {String} folderId
 */
function moveFile(fileId,folderId) {
  let folder = DriveApp.getFolderById(folderId);
  let file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
}