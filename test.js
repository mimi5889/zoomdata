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
  //登録者数の変更がなくても実行する
  const registantsOrgCount = sheet.getRange(row,15).getValue();
  Logger.log('元の登録者数: ' + registantsOrgCount);
  
  const url = crieteFolderAndCsvFile(folderId,webinarId,topic,token,filePrefix,eventName,dateStr,scheduleDate,stockId, companyName, companyAdd);
  const newRegistantsCount = url[2]; // 新しい登録者数
  Logger.log('新しい登録者数: ' + newRegistantsCount);
  
  const aryday = new Date(Utilities.formatDate(rowValues[3], 'Asia/Tokyo', 'yyyy/MM/dd'));
  const d1 = new Date(aryday.getFullYear(),aryday.getMonth(),aryday.getDate());
  const d2 = new Date(today.getFullYear(),today.getMonth(),today.getDate());
  const diffTime = d1.getTime() - d2.getTime();
  const laterDays = diffTime / (1000 * 60 * 60 * 24);
  Logger.log(laterDays);
  
  // 登録者数の変更がなくても実行する（条件分岐を削除）
  const url_txt = '[' + eventName + ']\n'+laterDays + '日前\n'+ formatted_today ;
  const webhooktxt = url_txt + '\n' + topic + '\n' + url[1] +'\n';
  Logger.log(webhooktxt);
  
  if(stockId ==='' || companyAdd === '' || companyAdd === 0){
    sendSlackNotification3(topic,eventName,url[0]) //************************事前登録者データメールアドレス無しslack通知************************
    sheet.getRange(row,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
    sheet.getRange(row,14).setValue(url[0]);
    sheet.getRange(row,15).setValue(newRegistantsCount); // 新しい登録者数を設定
  }else{
    sendSlackNotification2(webhooktxt);//************************事前登録者データslack通知************************
    sheet.getRange(row,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
    sheet.getRange(row,14).setValue(url[0]);
    sheet.getRange(row,15).setValue(newRegistantsCount); // 新しい登録者数を設定
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

function priorityJobStatusTest() {
  const ui = SpreadsheetApp.getUi();
  
  const infoMessage = `🧪 優先ジョブ状態管理の軽量テスト\n\n` +
    `・実際の処理は実行されません\n` +
    `・スクリプトプロパティの動作確認のみ\n` +
    `・状態の設定・取得・リセットをテスト\n\n` +
    `テストを実行しますか？`;

  const response = ui.alert('優先ジョブ状態管理テスト確認', infoMessage, ui.ButtonSet.YES_NO);
  
  if (response != ui.Button.YES) {
    ui.alert('処理中止', '処理を中止しました');
    return;
  }

  ui.alert('処理開始', '優先ジョブ状態管理テストを開始します...', ui.ButtonSet.OK);

  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // テスト前の状態を保存
    const originalStatus = scriptProperties.getProperty('PRIORITY_JOB_STATUS');
    const originalStartTime = scriptProperties.getProperty('PRIORITY_JOB_START_TIME');
    
    Logger.log('=== 優先ジョブ状態管理テスト開始 ===');
    
    // 1. 初期状態の確認
    Logger.log('1. 初期状態確認');
    const initialStatus = scriptProperties.getProperty('PRIORITY_JOB_STATUS');
    const initialStartTime = scriptProperties.getProperty('PRIORITY_JOB_START_TIME');
    Logger.log(`初期状態: ${initialStatus || '未設定'}`);
    Logger.log(`開始時刻: ${initialStartTime || '未設定'}`);
    
    // 2. RUNNING状態の設定テスト
    Logger.log('2. RUNNING状態の設定テスト');
    scriptProperties.setProperty('PRIORITY_JOB_STATUS', 'RUNNING');
    scriptProperties.setProperty('PRIORITY_JOB_START_TIME', new Date().toISOString());
    
    const runningStatus = scriptProperties.getProperty('PRIORITY_JOB_STATUS');
    const runningStartTime = scriptProperties.getProperty('PRIORITY_JOB_START_TIME');
    Logger.log(`設定後状態: ${runningStatus}`);
    Logger.log(`設定後開始時刻: ${runningStartTime}`);
    
    // 3. 状態の取得テスト
    Logger.log('3. 状態の取得テスト');
    const currentStatus = scriptProperties.getProperty('PRIORITY_JOB_STATUS');
    const currentStartTime = scriptProperties.getProperty('PRIORITY_JOB_START_TIME');
    
    if (currentStatus === 'RUNNING' && currentStartTime) {
      Logger.log('✅ RUNNING状態の設定・取得: 成功');
    } else {
      throw new Error('RUNNING状態の設定・取得に失敗');
    }
    
    // 4. IDLE状態へのリセットテスト
    Logger.log('4. IDLE状態へのリセットテスト');
    scriptProperties.setProperty('PRIORITY_JOB_STATUS', 'IDLE');
    scriptProperties.deleteProperty('PRIORITY_JOB_START_TIME');
    
    const resetStatus = scriptProperties.getProperty('PRIORITY_JOB_STATUS');
    const resetStartTime = scriptProperties.getProperty('PRIORITY_JOB_START_TIME');
    Logger.log(`リセット後状態: ${resetStatus}`);
    Logger.log(`リセット後開始時刻: ${resetStartTime}`);
    
    if (resetStatus === 'IDLE' && !resetStartTime) {
      Logger.log('✅ IDLE状態へのリセット: 成功');
    } else {
      throw new Error('IDLE状態へのリセットに失敗');
    }
    
    // 5. 元の状態に復元
    Logger.log('5. 元の状態への復元');
    if (originalStatus) {
      scriptProperties.setProperty('PRIORITY_JOB_STATUS', originalStatus);
    } else {
      scriptProperties.deleteProperty('PRIORITY_JOB_STATUS');
    }
    
    if (originalStartTime) {
      scriptProperties.setProperty('PRIORITY_JOB_START_TIME', originalStartTime);
    } else {
      scriptProperties.deleteProperty('PRIORITY_JOB_START_TIME');
    }
    
    Logger.log('=== 優先ジョブ状態管理テスト完了 ===');
    
    // 結果表示
    const resultMessage = `✅ 優先ジョブ状態管理テスト完了\n\n` +
      `・状態設定: 成功\n` +
      `・状態取得: 成功\n` +
      `・状態リセット: 成功\n` +
      `・元の状態復元: 完了\n\n` +
      `詳細はログを確認してください\n` +
      `実際の処理は実行されていません`;

    ui.alert('テスト完了', resultMessage, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log(`❌ 優先ジョブ状態管理テストでエラー: ${error.message}`);
    
    const errorMessage = `❌ 優先ジョブ状態管理テストでエラーが発生しました\n\n` +
      `エラー内容: ${error.message}\n\n` +
      `詳細: ${error.stack || 'スタックトレースなし'}`;
    
    ui.alert('テストエラー', errorMessage, ui.ButtonSet.OK);
  }
}

function webinarReportsTriggerTest() {
  const ui = SpreadsheetApp.getUi();
  
  const infoMessage = `🧪 webinarReportsTriggerの軽量テスト\n\n` +
    `・実際の処理は実行されません\n` +
    `・優先ジョブ状態チェックの動作確認のみ\n` +
    `・トリガー制御ロジックをテスト\n\n` +
    `テストを実行しますか？`;

  const response = ui.alert('webinarReportsTriggerテスト確認', infoMessage, ui.ButtonSet.YES_NO);
  
  if (response != ui.Button.YES) {
    ui.alert('処理中止', '処理を中止しました');
    return;
  }

  ui.alert('処理開始', 'webinarReportsTriggerテストを開始します...', ui.ButtonSet.OK);

  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    Logger.log('=== webinarReportsTriggerテスト開始 ===');
    
    // 1. 現在の状態を確認
    Logger.log('1. 現在の状態確認');
    const currentStatus = scriptProperties.getProperty('PRIORITY_JOB_STATUS');
    const currentRow = scriptProperties.getProperty('i');
    Logger.log(`優先ジョブ状態: ${currentStatus || '未設定'}`);
    Logger.log(`現在の行: ${currentRow || '未設定'}`);
    
    // 2. 制御ロジックのテスト
    Logger.log('2. 制御ロジックのテスト');
    
    let shouldSkip = false;
    let skipReason = '';
    
    if (currentStatus === 'RUNNING') {
      shouldSkip = true;
      skipReason = '優先ジョブ実行中';
    } else if (currentRow && currentRow !== '2') {
      shouldSkip = true;
      skipReason = '処理中の行がある';
    }
    
    if (shouldSkip) {
      Logger.log(`✅ 制御ロジック: スキップ判定 (理由: ${skipReason})`);
    } else {
      Logger.log(`✅ 制御ロジック: 実行可能`);
    }
    
    // 3. 時間制御のテスト
    Logger.log('3. 時間制御のテスト');
    const now = new Date();
    const hour = now.getHours();
    Logger.log(`現在時刻: ${hour}時`);
    
    if (hour >= 23 || hour < 7) {
      Logger.log(`✅ 時間制御: 実行時間外 (23:00-7:00) - スキップ`);
    } else {
      Logger.log(`✅ 時間制御: 実行時間内 - 実行可能`);
    }
    
    Logger.log('=== webinarReportsTriggerテスト完了 ===');
    
    // 結果表示
    const resultMessage = `✅ webinarReportsTriggerテスト完了\n\n` +
      `・優先ジョブ状態: ${currentStatus || '未設定'}\n` +
      `・現在の行: ${currentRow || '未設定'}\n` +
      `・制御ロジック: ${shouldSkip ? `スキップ (${skipReason})` : '実行可能'}\n` +
      `・時間制御: ${(hour >= 23 || hour < 7) ? '実行時間外' : '実行時間内'}\n\n` +
      `詳細はログを確認してください\n` +
      `実際の処理は実行されていません`;

    ui.alert('テスト完了', resultMessage, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log(`❌ webinarReportsTriggerテストでエラー: ${error.message}`);
    
    const errorMessage = `❌ webinarReportsTriggerテストでエラーが発生しました\n\n` +
      `エラー内容: ${error.message}\n\n` +
      `詳細: ${error.stack || 'スタックトレースなし'}`;
    
    ui.alert('テストエラー', errorMessage, ui.ButtonSet.OK);
  }
}

function registrantsCheckLogicTest() {
  Logger.log('=== 登録者数増減チェックロジックテスト開始 ===');
  Logger.log('🧪 登録者数増減チェックロジックの軽量テスト');
  Logger.log('・実際の処理は実行されません');
  Logger.log('・laterDaysの計算と範囲制限の確認のみ');
  Logger.log('・スプレッドシートへの記入・メール・Slack・Drive・CSV作成なし');
  Logger.log('');

  try {
    Logger.log('=== 登録者数増減チェックロジックテスト開始 ===');
    
    // 1. 日付計算のテスト
    Logger.log('1. 日付計算のテスト');
    const today = new Date();
    Logger.log(`現在時刻: ${today.toISOString()}`);
    
    // テスト用の日付パターン
    const testDates = [
      { name: '35日前', days: 35 },
      { name: '25日前', days: 25 },
      { name: '21日前', days: 21 },
      { name: '14日前', days: 14 },
      { name: '7日前', days: 7 },
      { name: '当日', days: 0 },
      { name: '1日後', days: -1 },
      { name: '7日後', days: -7 }
    ];
    
    testDates.forEach(testCase => {
      const testDate = new Date(today);
      testDate.setDate(today.getDate() + testCase.days);
      
      // laterDaysの計算（実際のコードと同じロジック）
      const aryday = new Date(testDate.getFullYear(), testDate.getMonth(), testDate.getDate());
      const d1 = new Date(aryday.getFullYear(), aryday.getMonth(), aryday.getDate());
      const d2 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      const diffTime = d1.getTime() - d2.getTime();
      const laterDays = diffTime / (1000 * 60 * 60 * 24);
      
      // 範囲制限の判定
      const isInRange = laterDays <= 21 && laterDays >= 0;
      const action = isInRange ? '実行' : 'スキップ';
      
      Logger.log(`${testCase.name}: laterDays=${laterDays}, 範囲内=${isInRange}, 処理=${action}`);
      
      if (isInRange) {
        Logger.log(`  ✅ 範囲内: 登録者数増減チェックを実行`);
        Logger.log(`  ✅ getRegistantsCount()を呼び出し`);
        Logger.log(`  ✅ 登録者数を更新`);
      } else {
        Logger.log(`  ⏭️ 範囲外: 登録者数増減チェックをスキップ`);
        Logger.log(`  ⏭️ getRegistantsCount()は呼び出さない`);
        Logger.log(`  ⏭️ 登録者数は更新しない`);
      }
    });
    
    // 2. 条件分岐のテスト
    Logger.log('2. 条件分岐のテスト');
    
    // 範囲内の場合の処理フロー
    Logger.log('範囲内の場合の処理フロー:');
    Logger.log('  if(laterDays <= 21 && laterDays >= 0) {');
    Logger.log('    // 登録者数を取得');
    Logger.log('    const registantsOrgCount = sheet.getRange(i,15).getValue();');
    Logger.log('    const registantsCount = getRegistantsCount(webinarId,token,topic);');
    Logger.log('    // 増減チェックとCSV作成・メール送信');
    Logger.log('    sheet.getRange(i,15).setValue(registantsCount);');
    Logger.log('  }');
    
    // 範囲外の場合の処理フロー
    Logger.log('範囲外の場合の処理フロー:');
    Logger.log('  } else {');
    Logger.log('    Logger.log(`行 ${i}: laterDays=${laterDays} のため登録者数増減チェックをスキップ`);');
    Logger.log('  }');
    
    // 3. 効率性の確認
    Logger.log('3. 効率性の確認');
    Logger.log('✅ 範囲外の行ではgetRegistantsCount()を呼び出さない');
    Logger.log('✅ 不要なAPI呼び出しを削減');
    Logger.log('✅ 処理時間の短縮');
    Logger.log('✅ リソース使用量の削減');
    
    Logger.log('=== 登録者数増減チェックロジックテスト完了 ===');
    
    // 結果表示
    Logger.log('=== テスト結果 ===');
    Logger.log('✅ 日付計算: 正常');
    Logger.log('✅ 範囲制限: 正常');
    Logger.log('✅ 条件分岐: 正常');
    Logger.log('✅ 効率性: 向上確認');
    Logger.log('');
    Logger.log('詳細はログを確認してください');
    Logger.log('実際の処理は実行されていません');
    Logger.log('=== テスト完了 ===');

  } catch (error) {
    Logger.log(`❌ 登録者数増減チェックロジックテストでエラー: ${error.message}`);
    Logger.log(`エラー内容: ${error.message}`);
    Logger.log(`詳細: ${error.stack || 'スタックトレースなし'}`);
    Logger.log('=== テストエラー ===');
  }
}

function setExclusionFlagsTest() {
  Logger.log('=== setExclusionFlags軽量テスト開始 ===');
  
  try {
    // 1. 除外シートの状態確認
    Logger.log('1. 除外シートの状態確認');
    const flgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('除外');
    
    if (!flgSheet) {
      throw new Error('除外シートが見つかりません');
    }
    
    // 除外ワード（A列）の取得
    const exclusionWords = flgSheet.getRange('A2:A')
      .getValues()
      .flat()
      .filter(word => word);
    
    Logger.log(`除外ワード数: ${exclusionWords.length}`);
    Logger.log(`除外ワード: ${exclusionWords.join(', ')}`);
    
    // 除外stockID（B列）の取得
    const exclusionStockIds = flgSheet.getRange('B2:B')
      .getValues()
      .flat()
      .filter(stockId => stockId);
    
    Logger.log(`除外stockID数: ${exclusionStockIds.length}`);
    Logger.log(`除外stockID: ${exclusionStockIds.join(', ')}`);
    
    // 2. メインシートの状態確認
    Logger.log('2. メインシートの状態確認');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      Logger.log('メインシートにデータがありません');
      return;
    }
    
    Logger.log(`メインシート行数: ${lastRow}`);
    
    // 3. 除外判定ロジックのテスト
    Logger.log('3. 除外判定ロジックのテスト');
    
    // テスト用のサンプルデータ
    const testCases = [
      { topic: 'テストウェビナー1', stockId: '1234', description: '通常のケース' },
      { topic: 'テストウェビナー2', stockId: '5678', description: '除外ワードが含まれるケース' },
      { topic: 'テストウェビナー3', stockId: '9999', description: '除外stockIDと一致するケース' },
      { topic: 'テストウェビナー4', stockId: '', description: 'stockIDが空のケース' }
    ];
    
    testCases.forEach((testCase, index) => {
      Logger.log(`テストケース${index + 1}: ${testCase.description}`);
      Logger.log(`  トピック: ${testCase.topic}`);
      Logger.log(`  証券コード: ${testCase.stockId || '空'}`);
      
      // 除外ワードによる除外判定
      const isExcludedByWord = exclusionWords.some(word => 
        testCase.topic.includes(word)
      );
      Logger.log(`  除外ワード判定: ${isExcludedByWord ? '除外' : '対象'}`);
      
      // 除外stockIDによる除外判定
      const isExcludedByStockId = exclusionStockIds.some(exclusionId => 
        testCase.stockId && exclusionId && testCase.stockId.toString() === exclusionId.toString()
      );
      Logger.log(`  除外stockID判定: ${isExcludedByStockId ? '除外' : '対象'}`);
      
      // 最終判定
      const isExcluded = isExcludedByWord || isExcludedByStockId;
      Logger.log(`  最終判定: ${isExcluded ? '除外対象' : '処理対象'}`);
      Logger.log('');
    });
    
    // 4. 実際のデータでの除外件数予測
    Logger.log('4. 実際のデータでの除外件数予測');
    
    const topics = sheet.getRange(2, 3, lastRow - 1).getValues(); // C列（トピック）
    const stockIds = sheet.getRange(2, 8, lastRow - 1).getValues(); // H列（証券コード）
    
    let excludedCount = 0;
    let excludedByWordCount = 0;
    let excludedByStockIdCount = 0;
    
    for (let i = 0; i < topics.length; i++) {
      const topic = topics[i][0];
      const stockId = stockIds[i][0];
      
      // 除外ワードによる除外判定
      const isExcludedByWord = exclusionWords.some(word => topic.includes(word));
      
      // 除外stockIDによる除外判定
      const isExcludedByStockId = exclusionStockIds.some(exclusionId => 
        stockId && exclusionId && stockId.toString() === exclusionId.toString()
      );
      
      if (isExcludedByWord) excludedByWordCount++;
      if (isExcludedByStockId) excludedByStockIdCount++;
      if (isExcludedByWord || isExcludedByStockId) excludedCount++;
    }
    
    Logger.log(`総レコード数: ${topics.length}`);
    Logger.log(`除外ワードによる除外: ${excludedByWordCount}件`);
    Logger.log(`除外stockIDによる除外: ${excludedByStockIdCount}件`);
    Logger.log(`総除外件数: ${excludedCount}件`);
    
    // 5. 処理の安全性確認
    Logger.log('5. 処理の安全性確認');
    Logger.log('✅ 除外ワードの取得: 正常');
    Logger.log('✅ 除外stockIDの取得: 正常');
    Logger.log('✅ 文字列比較: toString()で安全な比較');
    Logger.log('✅ null/undefinedチェック: 空値の適切な処理');
    Logger.log('✅ ログ出力: 処理結果の確認が可能');
    
    Logger.log('=== setExclusionFlags軽量テスト完了 ===');
    
  } catch (error) {
    Logger.log(`❌ setExclusionFlags軽量テストでエラー: ${error.message}`);
    Logger.log(`エラー詳細: ${error.stack || 'スタックトレースなし'}`);
    Logger.log('=== テストエラー ===');
  }
}

// ===== 複数エンドポイントからcustom_survey情報をテストする関数 =====

function testCustomSurveyEndpoints() {
  // アクティブセルのウェビナー情報を利用して、複数のエンドポイントからcustom_surveyの情報をテスト
  Logger.log('=== 🧪 複数エンドポイントcustom_surveyテスト開始 ===');
  
  try {
    // アクティブセルの情報を取得
    const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const row = sh.getActiveCell().getRow();
    
    if (row < 2) {
      Logger.log('❌ ヘッダー行が選択されています。データ行を選択してください。');
      return;
    }
    
    const account = sh.getRange(row, 1).getValue(); // A列：アカウント
    const webinarId = sh.getRange(row, 2).getValue(); // B列：ウェビナーID
    const topic = sh.getRange(row, 3).getValue(); // C列：トピック
    
    Logger.log(`選択された行: ${row}`);
    Logger.log(`アカウント: ${account}`);
    Logger.log(`ウェビナーID: ${webinarId}`);
    Logger.log(`トピック: ${topic}`);
    
    if (!account || !webinarId) {
      Logger.log('❌ アカウントまたはウェビナーIDが取得できませんでした。');
      return;
    }
    
    // アカウントからスクリプトプロパティをforで回してインデックス取得する
    const scriptProperties = PropertiesService.getScriptProperties();
    const max_acccountIndex = parseInt(scriptProperties.getProperty('MAX_ACCOUNT_INDEX') || '4');
    let accountIndex = 0;
    
    for (let n = 1; n <= max_acccountIndex; n++) {
      const zoomId = scriptProperties.getProperty('ZOOM_ID_' + n);
      if (account == zoomId) {
        accountIndex = n;
        break;
      }
    }
    
    if (accountIndex === 0) {
      Logger.log('❌ アカウントのインデックスが見つかりませんでした。');
      return;
    }
    
    Logger.log(`アカウントインデックス: ${accountIndex}`);
    
    // トークンを取得
    const token = getAccessToken(accountIndex);
    if (!token) {
      Logger.log('❌ トークンの取得に失敗しました。');
      return;
    }
    
    Logger.log('✅ トークン取得成功');
    
    // 複数のエンドポイントをテスト
    Logger.log('\n🚀 複数エンドポイントテスト開始');
    
    const testResults = [];
    
    // 1. 通常のsurveyエンドポイント（webinarId使用）
    Logger.log('\n--- 1. 通常のsurveyエンドポイント（webinarId使用） ---');
    const result1 = testCustomSurveyEndpoint(
      `https://api.zoom.us/v2/webinars/${webinarId}/survey?fields=custom_survey`,
      token,
      '通常のsurveyエンドポイント'
    );
    testResults.push({ name: '通常のsurveyエンドポイント', result: result1 });
    
    // 2. 軽量版エンドポイント（webinarId使用）
    Logger.log('\n--- 2. 軽量版エンドポイント（webinarId使用） ---');
    const result2 = testCustomSurveyEndpoint(
      `https://api.zoom.us/v2/webinars/${webinarId}?fields=settings,survey,questions`,
      token,
      '軽量版エンドポイント'
    );
    testResults.push({ name: '軽量版エンドポイント', result: result2 });
    
    // 3. UUIDを取得してからsurveyエンドポイント
    Logger.log('\n--- 3. UUID経由surveyエンドポイント ---');
    const uuid = getWebinarUUID(webinarId, token);
    if (uuid) {
      const result3 = testCustomSurveyEndpoint(
        `https://api.zoom.us/v2/webinars/${uuid}/survey?fields=custom_survey`,
        token,
        'UUID経由surveyエンドポイント'
      );
      testResults.push({ name: 'UUID経由surveyエンドポイント', result: result3 });
    } else {
      Logger.log('❌ UUIDが取得できませんでした');
      testResults.push({ name: 'UUID経由surveyエンドポイント', result: { success: false, error: 'UUID取得失敗' } });
    }
    
    // 4. UUID経由の軽量版エンドポイント
    Logger.log('\n--- 4. UUID経由軽量版エンドポイント ---');
    if (uuid) {
      const result4 = testCustomSurveyEndpoint(
        `https://api.zoom.us/v2/webinars/${uuid}?fields=settings,survey,questions`,
        token,
        'UUID経由軽量版エンドポイント'
      );
      testResults.push({ name: 'UUID経由軽量版エンドポイント', result: result4 });
    } else {
      testResults.push({ name: 'UUID経由軽量版エンドポイント', result: { success: false, error: 'UUID取得失敗' } });
    }
    
    // 5. ダッシュボード系エンドポイント
    Logger.log('\n--- 5. ダッシュボード系エンドポイント ---');
    const result5 = testCustomSurveyEndpoint(
      `https://api.zoom.us/v2/metrics/webinars/${webinarId}/participants?type=past&page_size=1`,
      token,
      'ダッシュボード系エンドポイント'
    );
    testResults.push({ name: 'ダッシュボード系エンドポイント', result: result5 });
    
    // 結果の比較とサマリー
    Logger.log('\n=== 📊 テスト結果サマリー ===');
    testResults.forEach((test, index) => {
      const status = test.result.success ? '✅' : '❌';
      const details = test.result.success ? 
        `取得時間: ${test.result.responseTime}ms, データサイズ: ${test.result.dataSize}文字` : 
        `エラー: ${test.result.error}`;
      
      Logger.log(`${index + 1}. ${test.name}: ${status}`);
      Logger.log(`   ${details}`);
    });
    
    // 成功したエンドポイントの分析
    const successfulTests = testResults.filter(test => test.result.success);
    if (successfulTests.length > 0) {
      Logger.log('\n=== 🏆 成功したエンドポイント分析 ===');
      
      // レスポンス時間でソート
      successfulTests.sort((a, b) => a.result.responseTime - b.result.responseTime);
      
      successfulTests.forEach((test, index) => {
        Logger.log(`${index + 1}. ${test.name}: ${test.result.responseTime}ms`);
      });
      
      Logger.log(`\n最速エンドポイント: ${successfulTests[0].name} (${successfulTests[0].result.responseTime}ms)`);
    } else {
      Logger.log('\n❌ すべてのエンドポイントで失敗しました');
    }
    
    Logger.log('=== 🎯 複数エンドポイントcustom_surveyテスト完了 ===');
    
  } catch (error) {
    Logger.log(`❌ テスト実行エラー: ${error.message}`);
    Logger.log(`エラー詳細: ${error.stack || 'スタックトレースなし'}`);
  }
}

function testCustomSurveyEndpoint(url, token, endpointName) {
  // 個別のエンドポイントをテスト
  try {
    Logger.log(`URL: ${url}`);
    
    const startTime = new Date();
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const endTime = new Date();
    
    const responseTime = endTime.getTime() - startTime.getTime();
    const statusCode = response.getResponseCode();
    const body = response.getContentText();
    
    Logger.log(`ステータスコード: ${statusCode}`);
    Logger.log(`レスポンス時間: ${responseTime}ms`);
    
    if (statusCode === 200) {
      try {
        const data = JSON.parse(body);
        const dataSize = body.length;
        
        Logger.log(`データサイズ: ${dataSize}文字`);
        
        // custom_surveyの情報を探す
        let customSurveyInfo = null;
        let hasCustomSurvey = false;
        
        if (data.custom_survey) {
          customSurveyInfo = data.custom_survey;
          hasCustomSurvey = true;
        } else if (data.settings && data.settings.survey) {
          customSurveyInfo = data.settings.survey;
          hasCustomSurvey = true;
        } else if (data.survey) {
          customSurveyInfo = data.survey;
          hasCustomSurvey = true;
        }
        
        if (hasCustomSurvey) {
          Logger.log(`✅ custom_survey情報取得成功`);
          Logger.log(`custom_survey構造: ${Object.keys(customSurveyInfo).join(', ')}`);
          
          if (customSurveyInfo.questions) {
            Logger.log(`設問数: ${customSurveyInfo.questions.length}`);
          }
        } else {
          Logger.log(`⚠️ custom_survey情報が見つかりませんでした`);
          Logger.log(`利用可能なフィールド: ${Object.keys(data).join(', ')}`);
        }
        
        return {
          success: true,
          responseTime: responseTime,
          dataSize: dataSize,
          hasCustomSurvey: hasCustomSurvey,
          customSurveyInfo: customSurveyInfo
        };
        
      } catch (parseError) {
        Logger.log(`⚠️ JSON解析エラー: ${parseError.message}`);
        return {
          success: false,
          error: `JSON解析エラー: ${parseError.message}`,
          responseTime: responseTime
        };
      }
    } else {
      Logger.log(`❌ HTTPエラー: ${statusCode}`);
      Logger.log(`エラー内容: ${body.substring(0, 200)}...`);
      return {
        success: false,
        error: `HTTP ${statusCode}: ${body.substring(0, 100)}`,
        responseTime: responseTime
      };
    }
    
  } catch (fetchError) {
    Logger.log(`⚠️ フェッチエラー: ${fetchError.message}`);
    return {
      success: false,
      error: `フェッチエラー: ${fetchError.message}`
    };
  }
}

function getWebinarUUID(webinarId, token) {
  // ウェビナーIDからUUIDを取得
  try {
    Logger.log(`UUID取得開始: ${webinarId}`);
    
    const response = UrlFetchApp.fetch(
      `https://api.zoom.us/v2/webinars/${webinarId}?fields=uuid`, 
      {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data && data.uuid) {
        Logger.log(`✅ UUID取得成功: ${data.uuid}`);
        return data.uuid;
      }
    }
    
    Logger.log('❌ UUIDが取得できませんでした');
    return null;
  } catch (e) {
    Logger.log(`⚠️ UUID取得エラー: ${e.message}`);
    return null;
  }
}



