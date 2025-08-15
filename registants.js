//事前登録者データ取得

function priorityJob_registants() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // 最大30秒待ってロック取得
  try {
    // 優先処理本体
    getRegistantsData();
  } finally {
    lock.releaseLock();
  }
}


function getRegistantsData() {//登録者データ取得

  const scriptProperties = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const folderId = scriptProperties.getProperty('FOLDER_ID');
  const max_acccountIndex = parseInt(scriptProperties.getProperty('MAX_ACCOUNT_INDEX') || '4');//設定されていない場合は4とする

  const today = new Date();
  const formatted_today = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  let ary = new Array();
  let slackAry = new Array();
  let webhooktxt = '';
  let nomailCount = 0;//メールアドレス取得できない企業カウント
  let errorCount = parseInt(scriptProperties.getProperty('errorCount') || '0');
  let currentRow = parseInt(scriptProperties.getProperty('i') || '2');
  let accountIndex = 0;

  try{
    for (let i = currentRow;  i < sheet.getLastRow() ; i++){
      try {
        ary = sheet.getRange(i, 1, 1, sheet.getMaxColumns()).getValues()[0];
        const account = ary[0];

        //アカウントからスクリプトプロパティをforで回してインデックス取得する
        for(let n = 1 ; n <= max_acccountIndex ; n++){
          const zoomId = scriptProperties.getProperty('ZOOM_ID_' + n);
          if(account == zoomId){
            accountIndex = n;
          }
        }
        // 右から見て最初に空でないセルのインデックスを取得
        const pIndex = 15; // P列のインデックス（1始まりで16列目 → 0始まりで15）
        let colIndex = 15;
        for (let col = ary.length; col >  pIndex; col--) {
          if (ary[col - 1] !== '') {
            colIndex = col;
          }
        }
        Logger.log(colIndex);

        const webinarId = ary[1];
        const scheduleDate = Utilities.formatDate(ary[3], 'Asia/Tokyo', 'yyyy/MM/dd');
        const scheduleDate0 = new Date(ary[3]);
        scheduleDate0.setHours(0,0,0,0);
        today.setHours(0,0,0,0);
        Logger.log('ary[3]:' + ary[3]);
        Logger.log('scheduleDate:' + scheduleDate);
        const stockId = ary[7];
        const companyName = ary[8];
        const companyAdd = ary[9];
        const topic = ary[2];
        const filePrefix = `${companyName}(${stockId})様`;
        const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd');
        if (ary[6] === 1) continue;//除外フラグはパスする
        if (scheduleDate0.getTime() < today.getTime()) continue;//開催日が過去のものはパスする
        Logger.log(i);
        Logger.log(companyName);
        Logger.log(scheduleDate);
        const token = getAccessToken(accountIndex); // トークン取得関数に index を渡す
        let laterDays = 0;

        let schedule = new Date(today);
        Logger.log('today.getDate()'+today.getDate());
        laterDays = 21;
        schedule.setDate(today.getDate() + laterDays);
        if (Utilities.formatDate(ary[3], 'Asia/Tokyo', 'yyyy/MM/dd') ===
          Utilities.formatDate(schedule, 'Asia/Tokyo', 'yyyy/MM/dd')){//3週間前
          const eventName = '3週間前';
          Logger.log(eventName);
          Logger.log(topic);
          Logger.log(scheduleDate);

          const url = crieteFolderAndCsvFile(folderId,webinarId,topic,token,filePrefix,eventName,dateStr,scheduleDate,stockId, companyName, companyAdd);//csv作成メール送信
          const url_txt = '[' + eventName + ']\n'+laterDays + '日前\n'+ formatted_today ;
          if(stockId ===''){
            sendSlackNotification3(topic,eventName,url[0]) //************************事前登録者データメールアドレス無しslack通知************************
            sheet.getRange(i,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
            sheet.getRange(i,14).setValue(url[0]);
            sheet.getRange(i,15).setValue(url[2]);
            nomailCount = nomailCount + 1;
          }else{
            slackAry.push([laterDays,schedule,topic,eventName,url[1]]);
            sheet.getRange(i,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
            sheet.getRange(i,14).setValue(url[0]);
            sheet.getRange(i,15).setValue(url[2]);
          }
        }
        schedule = new Date(today);
        laterDays = 7;
        schedule.setDate(today.getDate() + laterDays);
        if (Utilities.formatDate(ary[3], 'Asia/Tokyo', 'yyyy/MM/dd') ===
          Utilities.formatDate(schedule, 'Asia/Tokyo', 'yyyy/MM/dd')){//1週間前
          const eventName = '1週間前';
          Logger.log(eventName);
          Logger.log(topic);
          Logger.log(scheduleDate);
          const url = crieteFolderAndCsvFile(folderId,webinarId,topic,token,filePrefix,eventName,dateStr,scheduleDate,stockId, companyName, companyAdd);//csv作成メール送信
          const url_txt = '[' + eventName + ']\n'+laterDays + '日前\n'+ formatted_today ;
          if(stockId ==='' || companyAdd === '' || companyAdd === 0){
            sendSlackNotification3(topic,eventName,url[0]) //************************事前登録者データメールアドレス無しslack通知************************
            sheet.getRange(i,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
            sheet.getRange(i,14).setValue(url[0]);
            sheet.getRange(i,15).setValue(url[2]);
            nomailCount = nomailCount + 1;
          }else{
            slackAry.push([laterDays,schedule,topic,eventName,url[1]]);
            sheet.getRange(i,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
            sheet.getRange(i,14).setValue(url[0]);
            sheet.getRange(i,15).setValue(url[2]);
          }
        }
        schedule = new Date(today);
        laterDays = 0;
        schedule.setDate(today.getDate() + laterDays);
        if (Utilities.formatDate(ary[3], 'Asia/Tokyo', 'yyyy/MM/dd') ===
          Utilities.formatDate(schedule, 'Asia/Tokyo', 'yyyy/MM/dd')){//当日
          const eventName = '当日';
          Logger.log(eventName);
          Logger.log(topic);
          Logger.log(scheduleDate);
          const url = crieteFolderAndCsvFile(folderId,webinarId,topic,token,filePrefix,eventName,dateStr,scheduleDate,stockId, companyName, companyAdd);//csv作成メール送信
          const url_txt = '[' + eventName + ']\n'+laterDays + '日前\n'+ formatted_today ;
          if(stockId ==='' || companyAdd === '' || companyAdd === 0){
            sendSlackNotification3(topic,eventName,url[0]) //************************事前登録者データメールアドレス無しslack通知************************
            sheet.getRange(i,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
            sheet.getRange(i,14).setValue(url[0]);
            sheet.getRange(i,15).setValue(url[2]);
            nomailCount = nomailCount + 1;
          }else{
            slackAry.push([laterDays,schedule,topic,eventName,url[1]]);
            sheet.getRange(i,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
            sheet.getRange(i,14).setValue(url[0]);
            sheet.getRange(i,15).setValue(url[2]);
          }
        } 

        //登録者数を取得
        //登録者数に増減があったら動作する

        const registantsOrgCount = sheet.getRange(i,15).getValue();
        const eventName = '登録者数変更';
        const registantsCount = getRegistantsCount(webinarId,token,topic);
        sheet.getRange(i,15).setValue(registantsCount);
        const aryday = new Date(Utilities.formatDate(ary[3], 'Asia/Tokyo', 'yyyy/MM/dd'));
        const d1 = new Date(aryday.getFullYear(),aryday.getMonth(),aryday.getDate());
        const d2 = new Date(today.getFullYear(),today.getMonth(),today.getDate());
        const diffTime = d1.getTime() - d2.getTime();
        laterDays = diffTime / (1000 * 60 * 60 * 24);
        //laterDays =  Math.floor((aryday.getTime() - today.getTime())/ (1000 * 60 * 60 * 24));
        Logger.log(eventName);
        Logger.log(topic);
        Logger.log('laterDays:' + laterDays);
        schedule = new Date(today);
        schedule.setDate(today.getDate() + laterDays);
        if(registantsOrgCount !== registantsCount && laterDays <=21 && laterDays >=0 ){
          const url = crieteFolderAndCsvFile(folderId,webinarId,topic,token,filePrefix,eventName,dateStr,scheduleDate,stockId, companyName, companyAdd);//csv作成メール送信
          const url_txt = '[' + eventName + ']\n'+laterDays + '日前\n'+ formatted_today ;
          if(stockId ==='' || companyAdd === '' || companyAdd === 0){
            sendSlackNotification3(topic,eventName,url[0]) //************************事前登録者データメールアドレス無しslack通知************************
            sheet.getRange(i,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
            sheet.getRange(i,14).setValue(url[0]);
            sheet.getRange(i,15).setValue(url[2]);
            nomailCount = nomailCount + 1;
          }else{
            slackAry.push([laterDays,schedule,topic,eventName,url[1]]);
            sheet.getRange(i,colIndex+1).setFormula(`=HYPERLINK("${url[1]}", "${url_txt}")`);
            sheet.getRange(i,14).setValue(url[0]);
            sheet.getRange(i,15).setValue(url[2]);
          }

        }

        // 進捗を保存
        scriptProperties.setProperty('i', String(parseInt(i + 1, 10)));
        scriptProperties.setProperty('slackAry', JSON.stringify(slackAry));

        //強制的にエラー
        //if (i === 11) throw new Error("エラーテスト");
        //if (i === 13) throw new Error("エラーテスト");
        
      } catch (rowError) {
        // 個別の行処理でエラーが発生した場合
        Logger.log(`行 ${i} の処理でエラー: ${rowError.message}`);
        
        // Googleサーバーエラーの場合
        if (rowError.message.includes('Service temporarily unavailable') || 
            rowError.message.includes('Internal error') ||
            rowError.message.includes('Quota exceeded') ||
            rowError.message.includes('Rate Limit Exceeded')) {
          
          if (errorCount >= 3) {
            // 3回以上エラーが発生した場合は中止
            webhooktxt = `⚠️Googleサーバーエラーが3回発生しました\n処理を中止します\n現在の行: ${i}\nエラー: ${rowError.message}`;
            sendSlackNotification2(webhooktxt); //************************個別行エラー処理slack通知************************
            scriptProperties.setProperty('i', '2');
            scriptProperties.setProperty('slackAry', 'NaN');
            scriptProperties.setProperty('errorCount', '0');
            return;
          } else {
            // 再実行を試行
            webhooktxt = `⚠️Googleサーバーエラーが発生しました\n1分後に再実行します\n現在の行: ${i}\nエラー: ${rowError.message}\n再実行回数: ${errorCount + 1}/3`;
            sendSlackNotification2(webhooktxt); //************************個別行エラー処理slack通知************************
            scriptProperties.setProperty('errorCount', String(errorCount + 1));
            setRetryTrigger();
            return;
          }
        } else {
          // その他のエラーの場合は次の行に進む
          Logger.log(`行 ${i} をスキップして次の行に進みます`);
          continue;
        }
      }
    }
    //配列で復元
    let slackAry_props = scriptProperties.getProperty('slackAry');
    slackAry = JSON.parse(slackAry_props);
    //slack通知用に並べ替え
    slackAry.sort(function(a, b) {
    return a[0] - b[0];
    });
    Logger.log(slackAry);
    for (let i = 0 ; i < slackAry.length ; i++){
      let slackAry_date = new Date(slackAry[i][1]);
      if(i == 0){
        webhooktxt = '開催日：'+ Utilities.formatDate(slackAry_date, 'Asia/Tokyo', 'yyyy/MM/dd') + '(' + slackAry[i][3] + ')\n'
        + slackAry[i][2] + '\n' + slackAry[i][4] +'\n';
      }else if (slackAry[i-1][0] == slackAry[i][0]){
        webhooktxt = webhooktxt + slackAry[i][2] + '\n' + slackAry[i][4] +'\n';
      }else{
        webhooktxt = webhooktxt + '\n開催日：'+ Utilities.formatDate(slackAry_date, 'Asia/Tokyo', 'yyyy/MM/dd') + '(' + slackAry[i][3] + ')\n'
        + slackAry[i][2] + '\n' + slackAry[i][4] +'\n';
      }
    }
  }catch(e){
    const msg = e.message || '';
    if (msg.includes("Address unavailable")){
      // Zoom APIエラーの場合
      if (errorCount >= 3) {
        // 3回以上エラーが発生した場合は中止
        webhooktxt = `⚠️Zoom APIエラーが3回発生しました\n処理を中止します\n行番号: ${currentRow}\nアカウント: ${ary[0] || '不明'}\nトピック: ${ary[2] || '不明'}\nエラー: ${msg}`;
        sendSlackNotification2(webhooktxt); //************************エラー処理slack通知************************
        scriptProperties.setProperty('i', '2');
        scriptProperties.setProperty('slackAry', 'NaN');
        scriptProperties.setProperty('errorCount', '0');
        return;
      } else {
        // 再実行を試行
        webhooktxt = `⚠️Zoom APIエラーが発生しました\n1分後に再実行します\n行番号: ${currentRow}\nアカウント: ${ary[0] || '不明'}\nトピック: ${ary[2] || '不明'}\nエラー: ${msg}\n再実行回数: ${errorCount + 1}/3`;
        sendSlackNotification2(webhooktxt); //************************エラー処理slack通知************************
        scriptProperties.setProperty('errorCount', String(errorCount + 1));
        setRetryTrigger();
        return;
      }
    }else if (msg.includes("Service temporarily unavailable") || 
              msg.includes("Internal error") ||
              msg.includes("Quota exceeded") ||
              msg.includes("Rate Limit Exceeded")){
      // Googleサーバーエラーの場合
      if (errorCount >= 3) {
        // 3回以上エラーが発生した場合は中止
        webhooktxt = `⚠️Googleサーバーエラーが3回発生しました\n処理を中止します\n行番号: ${currentRow}\nアカウント: ${ary[0] || '不明'}\nトピック: ${ary[2] || '不明'}\nエラー: ${msg}`;
        sendSlackNotification2(webhooktxt); //************************エラー処理slack通知************************
        scriptProperties.setProperty('i', '2');
        scriptProperties.setProperty('slackAry', 'NaN');
        scriptProperties.setProperty('errorCount', '0');
        return;
      } else {
        // 再実行を試行
        webhooktxt = `⚠️Googleサーバーエラーが発生しました\n1分後に再実行します\n行番号: ${currentRow}\nアカウント: ${ary[0] || '不明'}\nトピック: ${ary[2] || '不明'}\nエラー: ${msg}\n再実行回数: ${errorCount + 1}/3`;
        sendSlackNotification2(webhooktxt); //************************エラー処理slack通知************************
        scriptProperties.setProperty('errorCount', String(errorCount + 1));
        setRetryTrigger();
        return;
      }
    }else{
      // その他のエラーの場合
      if (errorCount >= 3) {
        // 3回以上エラーが発生した場合は中止
        webhooktxt = `⚠️その他のエラーが3回発生しました\n処理を中止します\n行番号: ${currentRow}\nアカウント: ${ary[0] || '不明'}\nトピック: ${ary[2] || '不明'}\nエラー: ${msg}`;
        sendSlackNotification2(webhooktxt); //************************エラー処理slack通知************************
        scriptProperties.setProperty('i', '2');
        scriptProperties.setProperty('slackAry', 'NaN');
        scriptProperties.setProperty('errorCount', '0');
        return;
      } else {
        // 再実行を試行
        webhooktxt = `⚠️その他のエラーが発生しました\n1分後に再実行します\n行番号: ${currentRow}\nトピック: ${ary[2] || '不明'}\nエラー: ${msg}\n再実行回数: ${errorCount + 1}/3`;
        Logger.log(webhooktxt);
        sendSlackNotification2(webhooktxt); //************************エラー処理slack通知************************
        scriptProperties.setProperty('errorCount', String(errorCount + 1));
        setRetryTrigger();
        return;
      }
    }
  }

  Logger.log(webhooktxt);
  if(slackAry.length == 0 && nomailCount == 0){
    webhooktxt = 'ℹ️本日の事前データ取得通知はありません';
    sendSlackNotification2(webhooktxt); //************************処理完了slack通知************************
  }else if(slackAry.length > 0){
    sendSlackNotification2(webhooktxt); //************************処理完了slack通知************************
  }
  scriptProperties.setProperty('i', '2');
  scriptProperties.setProperty('slackAry', 'NaN');
  scriptProperties.setProperty('errorCount', '0');

  console.log('-----処理終了------');

}

function registantsCsvReport(registrants,topic) { // 登録者レポート
  const customTitlesSet = new Set();
  registrants.forEach(p => {
    (p.custom_questions || []).forEach(q => customTitlesSet.add(q.title));
  });
  const customTitles = Array.from(customTitlesSet);
  Logger.log(customTitles);
  const headers = [
     'トピック','名（登録）', '姓（登録）', 'メール', '登録時間',
      ...customTitles
  ];

  const rows = [];

  registrants.forEach(p => {

    const customAnswersMap = {};
    (p.custom_questions || []).forEach(q => {
      customAnswersMap[q.title] = q.value;
    });
    const customAnswers = customTitles.map(title => customAnswersMap[title] || '');
    const create_time = p.create_time ? Utilities.formatDate(new Date(p.create_time), 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss') : '-';

    rows.push([
      topic,
      p.first_name || '',
      p.last_name || '',
      p.email || '',
      create_time,
      ...customAnswers
    ]);
  });

  // rowsが空ならスキップ
  if (rows.length > 0) {
    const columnCount = headers.length;

    // 削除対象の列インデックスを格納
    const emptyColIndexes = [];

    for (let col = 0; col < columnCount; col++) {
      const isAllEmpty = rows.every(row => !row[col] || row[col].toString().trim() === '');
      if (isAllEmpty) {
        emptyColIndexes.push(col);
      }
    }

    if(hasDuplicates(headers)){//重複がある場合チェック
      // 降順にソートして後ろから削除（インデックスずれを防ぐため）
      //回答のない質問を削除する
      emptyColIndexes.sort((a, b) => b - a).forEach(index => {
      headers.splice(index, 1);
      rows.forEach(row => row.splice(index, 1));
      });
    }
  }


  return [headers, ...rows]
    .map(row => row.map(cell => `"${cell}"`).join(','))
    .join('\n');
}

function crieteFolderAndCsvFile(folderId,webinarId,topic,token,filePrefix,eventName,dateStr,scheduleDate,stockId, companyName, companyAdd){//フォルダ作成からCSV作成、slack通知まで
  // フォルダ作成
  Logger.log('------登録者レポート--------');
  Logger.log(eventName);
  Logger.log(filePrefix);
  const folder = getOrCreateFolderByName(folderId, topic)
  const folderUrl = folder.getUrl();
  //登録者データ取得
  const registrants = fetchZoomData(`https://api.zoom.us/v2/webinars/${webinarId}/registrants?page_size=1000`, token);
  const registrantsCsv = registantsCsvReport(registrants.registrants,topic);
  const newlineCount = (registrantsCsv.match(/\n/g) || []).length;
  Logger.log('改行:'+newlineCount);
  const bom = '\uFEFF'; // UTF-8 BOMを付けることでExcelで誤認しにくくする 
  const registrantsCsvWithBom = bom + registrantsCsv;
  Logger.log(registrantsCsvWithBom);
  const registrantsBlob = Utilities.newBlob(registrantsCsvWithBom, MimeType.CSV, `${filePrefix}-registrants Report_${dateStr}.csv`);
  const registrantsFile = folder.createFile(registrantsBlob);

  Utilities.sleep(1000);
  if(eventName !== 'テスト手動取得'){
    if(stockId !==''){
      createRegistantsMail(stockId, companyName, companyAdd, registrantsFile);//************メール送信******************
    }
  }
  const urlAry = [folderUrl,registrantsFile.getUrl(),newlineCount];
  return urlAry;
}

function getRegistantsCount(webinarId,token,topic){//登録者数をカウントする

  //登録者データ取得
  const registrants = fetchZoomData(`https://api.zoom.us/v2/webinars/${webinarId}/registrants?page_size=1000`, token);
  const registrantsCsv = registantsCsvReport(registrants.registrants,topic);
  const newlineCount = (registrantsCsv.match(/\n/g) || []).length;
  Logger.log('改行:'+newlineCount);// 改行数＝登録者数
  return newlineCount;
}

function hasDuplicates(arr) {//配列に重複があるかチェック　ある場合はtrueを返す
  return new Set(arr).size !== arr.length;
}

function setRetryTrigger() {//再実行トリガー
  // 1分後に再実行
  ScriptApp.newTrigger('getRegistantsData')
    .timeBased()
    .after(1 * 60 * 1000) // 1分後
    .create();
}
