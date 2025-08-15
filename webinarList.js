
function triggerSet(){//4アカウント分まとめてトリガーで実行　0：00～1：00
  const scriptProperties = PropertiesService.getScriptProperties();
  const max_acccountIndex = parseInt(scriptProperties.getProperty('MAX_ACCOUNT_INDEX') || '4');

  for(let accountIndex = 1 ; accountIndex <= max_acccountIndex ; accountIndex ++){
    try{
      getWebinarList(accountIndex);
    }catch(e){
      console.log('エラーを検知しました。');
      console.log('エラー内容：'+e.message);
    }
  }
  Utilities.sleep(1000);
  setExclusionFlags();//除外フラグの設定
  Utilities.sleep(1000);
  updateEmailInfoFromCode();//
}


function getWebinarList(accountIndex){//各IDから情報を取得しスプレッドシートに転記する
  const scriptProperties = PropertiesService.getScriptProperties();
  const zoomId = scriptProperties.getProperty('ZOOM_ID_' + accountIndex);
  const token = getAccessToken(accountIndex);
  const userId = getZoomUserId(token);//ユーザーIDの取得

  updateUpcomingWebinarsToSheet(token,userId,zoomId);

}

function updateUpcomingWebinarsToSheet(token,userId,zoomId) {//各ZoomIDからウェビナーリストの取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  const now = new Date();//現在時刻
  // SSシート内容を全取得
  const data = sheet.getDataRange().getValues(); // ヘッダー込み
  const rows = data.slice(1); // ヘッダー以外のデータ取得
  const idColMap = new Map(); //

  rows.forEach((row, i) => {
    idColMap.set(row[1].toString(), i + 2); // B列のID → 行番号（ヘッダー+1）を紐づけるmap作成
  });

  // Zoom APIで予定中ウェビナー取得
  const url = `https://api.zoom.us/v2/users/${userId}/webinars?page_size=500`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });
  const webinars = JSON.parse(res.getContentText()).webinars;//JSON形式に変換

  if (data.length === 0) {
    sheet.appendRow(['zoomID','ウェビナーID', 'トピック', '開始時刻', '終了時刻']);
  }

  //webinarsの詳細データ取得
  webinars.forEach(w => {
    const detailUrl = `https://api.zoom.us/v2/webinars/${w.id}`;
    const detailRes = UrlFetchApp.fetch(detailUrl, {
      headers: { Authorization: 'Bearer ' + token }
    });
    const detail = JSON.parse(detailRes.getContentText());
    const qa = detail.settings.question_and_answer.enable;//Q&Aにチェックがあるかどうか
    Logger.log(qa);
    const id = detail.id.toString();//ウェビナーID
    const start = new Date(detail.start_time);//開始時刻
    Logger.log('detail.topic' + detail.topic);
    const end = new Date(start.getTime() + (detail.duration || 0) * 60000);//終了時刻（開始時刻に実際の開催時間を加算して終了時刻とみなす）
    const formattedStart = formatDate(start);//yyyy-MM-dd HH:mm形式に変換
    const formattedEnd = formatDate(end);//yyyy-MM-dd HH:mm形式に変換

    if (idColMap.has(id)) {//B列のID検索して存在する場合に実行
      const rowIndex = idColMap.get(id);//行番号の取得
      const existingEnd = new Date(sheet.getRange(rowIndex, 5).getValue());//E列　終了予定時刻
      Logger.log('existingEnd' + existingEnd) ;
      if (existingEnd > now || existingEnd == 'Invalid Date') {//終了予定時刻が現在時刻より大きい場合　未来の開催とみなす
        // ウェビナー未終了なので上書き更新
        sheet.getRange(rowIndex, 3, 1, 3).setValues([
          [detail.topic, formattedStart, formattedEnd]
        ]);
        Logger.log('------データ更新-------');
        Logger.log(detail.topic);
      }
    } else if (start >= new Date()){// IDが存在しない場合新規データとして追記
      sheet.appendRow([zoomId,id, detail.topic, formattedStart, formattedEnd]);
      Logger.log('--------新規データ--------');
      Logger.log(detail.topic);
    }
    Utilities.sleep(200); // Zoom APIのレート制限対策
  });

  Logger.log(`${webinars.length} 件のウェビナーを処理しました`);
}


function formatDate(date) {//日付形式の変換
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}


function setExclusionFlags() {//処理を除外するレコードにフラグを立てる
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const flgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('除外');
  // 除外ワードの取得（除外シートA列）
  const exclusionWords = flgSheet.getRange('A2:A')
    .getValues()
    .flat()
    .filter(word => word); // 空でないものだけ

  const lastRow = sheet.getLastRow();
  const topics = sheet.getRange(2, 3, lastRow).getValues(); // C列（トピック）
  const flags = [];

  for (let i = 0; i < topics.length; i++) {
    const topic = topics[i][0];
    const isExcluded = exclusionWords.some(word => topic.includes(word));
    flags.push([isExcluded ? 1 : '']);
  }
  sheet.getRange(2, 7, flags.length).setValues(flags); // G列にフラグ出力
}

function runAllZoomChecks() {//グレーアウト処理
  const allExistingIdsMap = {}; // zoomIdごとのIDリストを格納

  for (let i = 1; i <= 4; i++) {
    try{
      const token = getAccessToken(i);
      const scriptProperties = PropertiesService.getScriptProperties();
      const zoomId = scriptProperties.getProperty('ZOOM_ID_' + i);
      const userId = getZoomUserId(token);

      const url = `https://api.zoom.us/v2/users/${userId}/webinars?page_size=200`;
      const res = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token }
      });
      const webinars = JSON.parse(res.getContentText()).webinars || [];

      // Mapに格納（Setで保持）
      allExistingIdsMap[zoomId] = new Set(webinars.map(w => w.id.toString()));

      // 任意：各Zoom IDに対するシート更新処理などもここに書けます
      Utilities.sleep(200);

    }catch(e){
      Logger.log(`例外発生： - ${e.message}`);
    }
  }

  // ✅ グレーアウト処理を実行
  grayOutMissingWebinars(allExistingIdsMap);

}

function grayOutMissingWebinars(allExistingIdsMap) {//グレーアウト処理
   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  const data = sheet.getDataRange().getValues(); // 全データ
  const rows = data.slice(1); // ヘッダー除く

  rows.forEach((row, i) => {
    const zoomId = row[0];
    const webinarId = row[1]?.toString();
    const rowIndex = i + 2;

    if (!zoomId || !webinarId) return;

    const existingSet = allExistingIdsMap[zoomId];
    if (!existingSet || !existingSet.has(webinarId)) {
      sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn())
        .setBackground('#dddddd');
    } else {
      // 存在する場合は背景を白に戻す
      sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn())
        .setBackground(null);
    }
  });
}





