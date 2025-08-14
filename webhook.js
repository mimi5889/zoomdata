function sendSlackNotification(topic,folderUrl,webhooktxt,stockId,companyAdd) {//事後データslack通知
  const scriptProperties = PropertiesService.getScriptProperties();
  const webhookUrl = scriptProperties.getProperty('SLACK_WEBHOOK_URL');
  const sheetId = scriptProperties.getProperty('SHEET_ID');
  const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  const flgSheet = sheet.getSheetByName('除外');
  const exclusionIds = flgSheet.getRange('B2:B')//メールの自動送信を除外する証券コード
  .getValues()
  .flat()
  .filter(word => word); // 空でないものだけ
  let message = "";
  Logger.log('stockId' + stockId);
  if(exclusionIds.includes(stockId)){//除外IDと同一の場合
    message = {
      text: '【ウェビナー事後データ作成通知】\n'
      + '手動対応対象企業'+ '\n'
      + 'イベント名：' + topic + '\n'
      + webhooktxt + '\n'
      + '\nレポートフォルダ：' + folderUrl,
      username: "webinarデータ抽出",         // 任意の表示名
    };
  }else if (stockId === ''|| companyAdd ==='' || companyAdd === 0){//ストックIDがない場合
    message = {
      text: '【ウェビナー事後データ作成通知】\n'
      + '企業メールが取得できないため、メールを送付できません' + '\n'
      + 'イベント名：' + topic + '\n'
      + webhooktxt + '\n'
      + '\nレポートフォルダ：' + folderUrl,
      username: "webinarデータ抽出",         // 任意の表示名
    };
  }else{
    message = {
      text: '【ウェビナー事後データ作成通知】\n'
      + 'メールを送付しました'+ '\n'
      + 'イベント名：' + topic + '\n'
      + webhooktxt + '\n'
      + '\n※参考　レポートフォルダ：' + folderUrl,
      username: "webinarデータ抽出",         // 任意の表示名
    };

  }

  Logger.log(message);
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message)
  };
  Logger.log('slack通知前');
  UrlFetchApp.fetch(webhookUrl, options);
  Logger.log('slack通知後');
}


function sendSlackNotification2(webhooktxt) {//事前登録者データslack通知
  const scriptProperties = PropertiesService.getScriptProperties();
  const webhookUrl = scriptProperties.getProperty('SLACK_WEBHOOK_URL');

  const message = {
    text: '【ウェビナー事前登録者データ】\n'//***************
    + webhooktxt 
    ,
    username: "webinarデータ抽出",         // 任意の表示名
  };
  Logger.log(message);
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message)
  };

  UrlFetchApp.fetch(webhookUrl, options);
}

function sendSlackNotification3(topic,eventName,url) {//事前登録者データメールアドレス無しslack通知
  const scriptProperties = PropertiesService.getScriptProperties();
  const webhookUrl = scriptProperties.getProperty('SLACK_WEBHOOK_URL');

  const message = {
      text: '【ウェビナー事前データ作成通知】\n'
      + '企業メールが取得できないため、メールを送付できません' + '\n'
      + 'イベント名：' + topic + '（'+ eventName + '）'+'\n'
      + '\nレポートフォルダ：' + url,
      username: "webinarデータ抽出",         // 任意の表示名
  };
  Logger.log(message);

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message)
  };

  UrlFetchApp.fetch(webhookUrl, options);
}
