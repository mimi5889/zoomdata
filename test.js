function open() {
  var ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
  ui.createMenu('GAS実行')
  .addItem('事前参加者リスト手動実行', 'registantsTest')
  .addItem('事後データ取得', 'test')
  .addToUi();
}

function registantsTest(){//事前データテスト用
  const today = new Date();
  const formatted_today = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sh.getActiveCell().getRow();//選択セルの行を取得
  const account = sh.getRange(row,1).getValue();
  const scriptProperties = PropertiesService.getScriptProperties();

  //アカウントからスクリプトプロパティをforで回してインデックス取得する
  for(let n = 1 ; n <= max_acccountIndex ; n++){
    const zoomId = scriptProperties.getProperty('ZOOM_ID_' + n);
    if(account == zoomId){
      accountIndex = n;
    }
  }
  const rowValues = sheet.getRange(row, 1, 1, sheet.getMaxColumns()).getValues()[0];
  // 右から見て最初に空でないセルのインデックスを取得
  const pIndex = 15; // P列のインデックス（1始まりで16列目 → 0始まりで15）
  let colIndex = 15;
  for (let col = rowValues.length; col >  pIndex; col--) {
    if (rowValues[col - 1] !== '') {
      colIndex = col;
    }
  }
  Logger.log(colIndex);

  const webinarId  = sh.getRange(row,2).getValue();
  const topic  = sh.getRange(row,3).getValue();
  const startDate = sh.getRange(row,4).getValue();
  const stockId  = sh.getRange(row,8).getValue();
  const companyName = sh.getRange(row,9).getValue();
  const companyAdd = sh.getRange(row,10).getValue();
  const endTime = sh.getRange(row,6).getValue();
  const token = getAccessToken(accountIndex);
  const filePrefix = `${companyName}(${stockId})様`;
  const eventName = 'テスト手動取得';
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd');
  const scheduleDate = Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy/MM/dd');

  const ui = SpreadsheetApp.getUi();
    // 確認ダイアログを表示
  const response = ui.alert(
    '確認',
    topic + '\n'+
    'この処理を実行しますか？',
    ui.ButtonSet.YES_NO
  );

  // No を選んだらスクリプトを終了
  if (response != ui.Button.YES) {
    ui.alert('処理を中止しました');
    return; // ここで関数終了
  }

  //登録者数を取得
  //登録者数に増減があったら動作する
  const registantsOrgCount = sheet.getRange(row,15).getValue();
  Logger.log('登録者数' + registantsOrgCount);
  const url = crieteFolderAndCsvFile(folderId,webinarId,topic,token,filePrefix,eventName,dateStr,scheduleDate,stockId, companyName, companyAdd);
  const aryday = new Date(Utilities.formatDate(rowValues[3], 'Asia/Tokyo', 'yyyy/MM/dd'));
  const d1 = new Date(aryday.getFullYear(),aryday.getMonth(),aryday.getDate());
  const d2 = new Date(today.getFullYear(),today.getMonth(),today.getDate());
  const diffTime = d1.getTime() - d2.getTime();
  const laterDays = diffTime / (1000 * 60 * 60 * 24);
  //const laterDays =  Math.floor((aryday.getTime() - today.getTime())/ (1000 * 60 * 60 * 24));//****動くか確認
  Logger.log(laterDays);
  if(registantsOrgCount !== url[2]){
    const url_txt = '[' + eventName + ']\n'+laterDays + '日前\n'+ formatted_today ;
    const webhooktxt = url_txt + '\n' + topic + '\n' + url[1] +'\n';
    Logger.log(webhooktxt);
    if(stockId ==='' || companyAdd === '' || companyAdd === 0){
      sendSlackNotification3(topic,eventName,url[0]) //********事前登録者データメールアドレス無しslack通知*********
      sheet.getRange(row,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
      sheet.getRange(row,14).setValue(url[0]);
      sheet.getRange(row,15).setValue(url[2]);
    }else{
      sendSlackNotification2(webhooktxt);//slack通知
      sheet.getRange(row,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
      sheet.getRange(row,14).setValue(url[0]);
      sheet.getRange(row,15).setValue(url[2]);
    }

  }

}


function test(){//事後データテスト用
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sh.getActiveCell().getRow();//選択セルの行を取得
  const webinarId  = sh.getRange(row,2).getValue();
  const account = sh.getRange(row,1).getValue();
  const scriptProperties = PropertiesService.getScriptProperties();

  //アカウントからスクリプトプロパティをforで回してインデックス取得する
  for(let n = 1 ; n <= max_acccountIndex ; n++){
    const zoomId = scriptProperties.getProperty('ZOOM_ID_' + n);
    if(account == zoomId){
      accountIndex = n;
    }
  }
  const stockId  = sh.getRange(row,8).getValue();
  const companyName = sh.getRange(row,9).getValue();
  const companyAdd = sh.getRange(row,10).getValue();
  const endTime = sh.getRange(row,6).getValue();

  const result = exportWebinarCsvs(webinarId,  accountIndex ,stockId, companyName, endTime,companyAdd);

  if(!exclusionIds.includes(stockId)){
    createDraftMail(stockId,companyName,companyAdd,result.attendeeFile,result.surveyFile,result.qaFile);//****下書きメール作成****
  }

  sheet.getRange(row,11).setValue(result.fileUrls[0]);
  sheet.getRange(row,12).setValue(result.fileUrls[1]);
  sheet.getRange(row,13).setValue(result.fileUrls[2]);


}


