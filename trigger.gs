function setTrigger(){
  //毎週月曜の午前9~10時に繰り返し定期実行するトリガーを作成
  ScriptApp.newTrigger('updateData').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  
}