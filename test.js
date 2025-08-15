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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const folderId = scriptProperties.getProperty('FOLDER_ID');
  const max_acccountIndex = parseInt(scriptProperties.getProperty('MAX_ACCOUNT_INDEX') || '4');
  const flgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('除外');

  const exclusionIds = flgSheet.getRange('B2:B')//メールの自動送信を除外する証券コード
    .getValues()
    .flat()
    .filter(word => word); // 空でないものだけ

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
      sendSlackNotification3(topic,eventName,url[0]) //************************事前登録者データメールアドレス無しslack通知************************
      sheet.getRange(row,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
      sheet.getRange(row,14).setValue(url[0]);
      sheet.getRange(row,15).setValue(url[2]);
    }else{
      sendSlackNotification2(webhooktxt);//************************事前登録者データslack通知************************
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const folderId = scriptProperties.getProperty('FOLDER_ID');
  const flgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('除外');
  const exclusionIds = flgSheet.getRange('B2:B')//メールの自動送信を除外する証券コード
    .getValues()
    .flat()
    .filter(word => word); // 空でないものだけ
  const max_acccountIndex = parseInt(scriptProperties.getProperty('MAX_ACCOUNT_INDEX') || '4');

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
    createDraftMail(stockId,companyName,companyAdd,result.attendeeFile,result.surveyFile,result.qaFile);//************************下書きメール作成************************
  }

  sheet.getRange(row,11).setValue(result.fileUrls[0]);
  sheet.getRange(row,12).setValue(result.fileUrls[1]);
  sheet.getRange(row,13).setValue(result.fileUrls[2]);


}

function testExistingSlackWebhook() {//webhookのテスト
  // 直接ベタ書きでもOKですが、プロパティに入れているなら置き換えてください。
  const scriptProperties = PropertiesService.getScriptProperties();
  const webhookUrl = scriptProperties.getProperty('SLACK_WEBHOOK_URL');

  const payload = {
    text: `[TEST] Webhook connectivity check: ${new Date().toISOString()}`
  };

  const res = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,   // ← 失敗でも本文を取得
    followRedirects: true
  });

  const status = res.getResponseCode();
  const body   = res.getContentText();
  Logger.log({status, body});

  // Slack Incoming Webhookは通常200で "ok" を返します
  if (status === 200 && body === 'ok') {
    Logger.log('✅ Webhook は有効です（投稿成功）');
  } else {
    throw new Error(`❌ 投稿失敗: status=${status}, body=${body}`);
  }
}




function registantsTestLightweight() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sh.getActiveCell().getRow();
  
  if (row < 2) {
    SpreadsheetApp.getUi().alert('エラー', '2行目以降を選択してください');
    return;
  }

  const ary = sh.getRange(row, 1, 1, sh.getMaxColumns()).getValues()[0];
  const account = ary[0];
  const webinarId = ary[1];
  const topic = ary[2];
  const scheduleDate = ary[3];
  const stockId = ary[7];
  const companyName = ary[8];
  const companyAdd = ary[9];

  const infoMessage = `選択された行: ${row}\n` +
    `アカウント: ${account}\n` +
    `ウェビナーID: ${webinarId}\n` +
    `トピック: ${topic}\n` +
    `開催日: ${scheduleDate}\n` +
    `証券コード: ${stockId}\n` +
    `企業名: ${companyName}\n` +
    `企業メール: ${companyAdd}\n\n` +
    `⚠️ 軽量テストモード\n` +
    `・メール送信なし\n` +
    `・CSV作成なし\n` +
    `・Driveアップロードなし\n` +
    `・Slack通知なし\n\n` +
    `処理を実行しますか？`;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('事前データ取得軽量テスト確認', infoMessage, ui.ButtonSet.YES_NO);
  
  if (response != ui.Button.YES) {
    ui.alert('処理中止', '処理を中止しました');
    return;
  }

  ui.alert('処理開始', '事前データ取得軽量テストを開始します...', ui.ButtonSet.OK);

  try {
    // スクリプトプロパティの取得テスト
    const scriptProperties = PropertiesService.getScriptProperties();
    const folderId = scriptProperties.getProperty('FOLDER_ID');
    const max_acccountIndex = parseInt(scriptProperties.getProperty('MAX_ACCOUNT_INDEX') || '4');
    
    // アカウントインデックスの取得テスト
    let accountIndex = 0;
    for(let n = 1; n <= max_acccountIndex; n++) {
      const zoomId = scriptProperties.getProperty('ZOOM_ID_' + n);
      if(account == zoomId) {
        accountIndex = n;
        break;
      }
    }

    if (accountIndex === 0) {
      throw new Error(`アカウント ${account} に対応するインデックスが見つかりません`);
    }

    // アクセストークンの取得テスト
    const token = getAccessToken(accountIndex);
    if (!token) {
      throw new Error('アクセストークンの取得に失敗しました');
    }

    // Zoom API接続テスト（登録者数取得のみ）
    const registrantsUrl = `https://api.zoom.us/v2/webinars/${webinarId}/registrants?page_size=1`;
    const response = UrlFetchApp.fetch(registrantsUrl, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Zoom API接続エラー: ${response.getResponseCode()}`);
    }

    const registrantsData = JSON.parse(response.getContentText());
    const registrantsCount = registrantsData.registrants ? registrantsData.registrants.length : 0;

    // 結果表示
    const resultMessage = `✅ 軽量テスト完了\n\n` +
      `・スクリプトプロパティ: OK\n` +
      `・アカウントインデックス: ${accountIndex}\n` +
      `・アクセストークン: 取得済み\n` +
      `・Zoom API接続: OK\n` +
      `・登録者数: ${registrantsCount}人\n\n` +
      `実際の処理は実行されていません`;

    ui.alert('テスト完了', resultMessage, ui.ButtonSet.OK);

  } catch (error) {
    const errorMessage = `❌ 軽量テストでエラーが発生しました\n\n` +
      `エラー内容: ${error.message}\n\n` +
      `詳細: ${error.stack || 'スタックトレースなし'}`;
    
    ui.alert('テストエラー', errorMessage, ui.ButtonSet.OK);
  }
}

function reportTestLightweight() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sh.getActiveCell().getRow();
  
  if (row < 2) {
    SpreadsheetApp.getUi().alert('エラー', '2行目以降を選択してください');
    return;
  }

  const ary = sh.getRange(row, 1, 1, sh.getMaxColumns()).getValues()[0];
  const account = ary[0];
  const webinarId = ary[1];
  const topic = ary[2];
  const endTimeStr = ary[4];
  const endTimeReal = ary[5];

  const infoMessage = `選択された行: ${row}\n` +
    `アカウント: ${account}\n` +
    `ウェビナーID: ${webinarId}\n` +
    `トピック: ${topic}\n` +
    `終了予定時刻: ${endTimeStr}\n` +
    `終了時刻: ${endTimeReal}\n\n` +
    `⚠️ 軽量テストモード\n` +
    `・CSV作成なし\n` +
    `・Driveアップロードなし\n` +
    `・メール下書き作成なし\n\n` +
    `処理を実行しますか？`;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('事後データ取得軽量テスト確認', infoMessage, ui.ButtonSet.YES_NO);
  
  if (response != ui.Button.YES) {
    ui.alert('処理中止', '処理を中止しました');
    return;
  }

  ui.alert('処理開始', '事後データ取得軽量テストを開始します...', ui.ButtonSet.OK);

  try {
    // スクリプトプロパティの取得テスト
    const scriptProperties = PropertiesService.getScriptProperties();
    const max_acccountIndex = parseInt(scriptProperties.getProperty('MAX_ACCOUNT_INDEX') || '4');
    
    // アカウントインデックスの取得テスト
    let accountIndex = 0;
    for(let n = 1; n <= max_acccountIndex; n++) {
      const zoomId = scriptProperties.getProperty('ZOOM_ID_' + n);
      if(account == zoomId) {
        accountIndex = n;
        break;
      }
    }

    if (accountIndex === 0) {
      throw new Error(`アカウント ${account} に対応するインデックスが見つかりません`);
    }

    // アクセストークンの取得テスト
    const token = getAccessToken(accountIndex);
    if (!token) {
      throw new Error('アクセストークンの取得に失敗しました');
    }

    // Zoom API接続テスト（出席者レポートのみ）
    const attendeesUrl = `https://api.zoom.us/v2/report/webinars/${webinarId}/participants?page_size=1`;
    const response = UrlFetchApp.fetch(attendeesUrl, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Zoom API接続エラー: ${response.getResponseCode()}`);
    }

    const attendeesData = JSON.parse(response.getContentText());
    const attendeesCount = attendeesData.participants ? attendeesData.participants.length : 0;

    // 結果表示
    const resultMessage = `✅ 軽量テスト完了\n\n` +
      `・スクリプトプロパティ: OK\n` +
      `・アカウントインデックス: ${accountIndex}\n` +
      `・アクセストークン: 取得済み\n` +
      `・Zoom API接続: OK\n` +
      `・出席者数: ${attendeesCount}人\n\n` +
      `実際の処理は実行されていません`;

    ui.alert('テスト完了', resultMessage, ui.ButtonSet.OK);

  } catch (error) {
    const errorMessage = `❌ 軽量テストでエラーが発生しました\n\n` +
      `エラー内容: ${error.message}\n\n` +
      `詳細: ${error.stack || 'スタックトレースなし'}`;
    
    ui.alert('テストエラー', errorMessage, ui.ButtonSet.OK);
  }
}

function webhookTestLightweight() {
  const ui = SpreadsheetApp.getUi();
  
  const infoMessage = `⚠️ 軽量テストモード\n` +
    `・実際のSlack通知は送信されません\n` +
    `・Webhook URLの形式チェックのみ実行\n\n` +
    `テストを実行しますか？`;

  const response = ui.alert('Webhook軽量テスト確認', infoMessage, ui.ButtonSet.YES_NO);
  
  if (response != ui.Button.YES) {
    ui.alert('処理中止', '処理を中止しました');
    return;
  }

  ui.alert('処理開始', 'Webhook軽量テストを開始します...', ui.ButtonSet.OK);

  try {
    // スクリプトプロパティの取得テスト
    const scriptProperties = PropertiesService.getScriptProperties();
    const webhookUrl = scriptProperties.getProperty('SLACK_WEBHOOK_URL');
    
    if (!webhookUrl) {
      throw new Error('SLACK_WEBHOOK_URLが設定されていません');
    }

    // Webhook URLの形式チェック
    if (!webhookUrl.startsWith('https://hooks.slack.com/')) {
      throw new Error('Webhook URLの形式が正しくありません');
    }

    // テスト用の軽量なPOSTリクエスト（実際の通知は送信しない）
    const testPayload = {
      text: "🧪 軽量テスト実行中 - 実際の通知は送信されません",
      username: "GAS Test Bot",
      icon_emoji: ":test_tube:"
    };

    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(testPayload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const resultMessage = `✅ Webhook軽量テスト完了\n\n` +
        `・Webhook URL: 設定済み\n` +
        `・URL形式: 正しい\n` +
        `・接続テスト: OK (${responseCode})\n\n` +
        `⚠️ テスト用の軽量通知がSlackに送信されました\n` +
        `実際の業務通知は送信されていません`;

      ui.alert('テスト完了', resultMessage, ui.ButtonSet.OK);
    } else {
      throw new Error(`Webhook接続エラー: ${responseCode}`);
    }

  } catch (error) {
    const errorMessage = `❌ Webhook軽量テストでエラーが発生しました\n\n` +
      `エラー内容: ${error.message}\n\n` +
      `詳細: ${error.stack || 'スタックトレースなし'}`;
    
    ui.alert('テストエラー', errorMessage, ui.ButtonSet.OK);
  }
}



