function checkText(subject,body) {//文字列チェックを行い含まれる場合のみ実行する
  // チェックする複数の文字列を定義
  const requiredTexts = [
    "説明会",
    "決算", 
    "投資", 
    "会場", 
    "IR", 
    "証券", 
    "カンファレンス", 
    "中期", 
    "コンファレンス", 
    "webiner", 
    "zoom"
    ];

  // 件名と本文の両方に特定の文字列が含まれるかをチェック
  const subjectMatches = requiredTexts.some(text => subject.includes(text));
  const bodyMatches = requiredTexts.some(text => body.includes(text));

  if (subjectMatches || bodyMatches) {
    // 条件を満たす場合に処理を実行
    console.log(`条件に一致するメールを検出しました。件名: ${subject}`);
    console.log(`本文: ${body}`);
    return true;//文字列が存在する 
  }
  return false;//文字列が存在しない
}

function checkCcAdd(ccAdd) {//ccチェックをおこない含まれる場合は除外する

  const excludeEmail = "investor@msetsu.com"; // CCに含まれるべきでないアドレス  
  
  if(ccAdd.includes(excludeEmail)){
    return true;//アドレスが存在する 
  }
  return false;//アドレスが存在しない
}

function getMail() {
  const userEmail = Session.getActiveUser().getEmail();
  //対象のスプレッドシートのIDを指定する
  const ss = SpreadsheetApp.openById('19a4KmBi2yXsLeCHz6N43Qjli5xXEuvqESTHmiSYUbmo');
  //書き込むシート名を指定する
  const sh = ss.getSheetByName("書き込み用");
  const startTime = new Date().getTime();
  const scriptProperties = PropertiesService.getScriptProperties();
  const propertyKey = "threadMessageIds";
  // スクリプトプロパティに保存されている以前のデータを取得
  const existingData = scriptProperties.getProperty(propertyKey);
  let data = existingData ? JSON.parse(existingData) : [];
  const newData = [];//スレッドIDとメッセージIDを入れる配列

  const url = scriptProperties.getProperty("API_URL");

  // 「処理済」ラベルを取得、なければ作成
  const processedLabel = GmailApp.getUserLabelByName('処理済') || GmailApp.createLabel('処理済');

  // 前回処理したメールの受信日時を取得（初回は当日0:00JSTを設定）
  let lastProcessedMessageDate = scriptProperties.getProperty("lastProcessedMessageDate");
  if (!lastProcessedMessageDate) {
    const todayStartJST = new Date();
    todayStartJST.setHours(0, 0, 0, 0); // JSTの0:00を設定
    lastProcessedMessageDate = todayStartJST.toISOString(); // 初回はJSTの0:00
  }

  console.log(`前回処理したメールの受信日時（JST）: ${convertToJST(new Date(lastProcessedMessageDate)).toISOString()}`);

  // Gmailからスレッドを取得（指定日時以降のメール、新しい順に取得される）
  const threads = GmailApp.search(`after:${Math.floor(new Date(lastProcessedMessageDate).getTime() / 1000)}`);
  console.log(`取得したスレッド数: ${threads.length}`);

  // スレッドを逆順に処理（古いスレッドから順に処理するため）
  let historyTime = new Date().getTime();
  for (let i = threads.length - 1; i >= 0; i--) {
    if (isTimeout(startTime)) {
      console.log("処理時間が制限に近いため中断します...");
      saveLastProcessedDate(scriptProperties, threads[i].getMessages());
      return;
    }
    let endTime = new Date().getTime();
    let elapsedSeconds = (endTime - startTime) / 1000;
    let historySeconds = (endTime - historyTime) / 1000;
    Logger.log(`経過時間: ${elapsedSeconds} 秒　　前回からの時間：${historySeconds} 秒`);
    historyTime = endTime;

    const thread = threads[i];

    // 最後のメッセージのみ取得（例: スレッド全体ではなく最終メールに限定）
    const messages = thread.getMessages();
    for (const message of messages) {
      if (isTimeout(startTime)) {
        console.log("処理時間が制限に近いため中断します...");
        saveLastProcessedDate(scriptProperties, messages);
        return;
      }

      const messageDate = convertToJST(message.getDate()); // JSTに変換

      // 最後に処理したメールの日時より新しいメールのみを処理
      if (messageDate > convertToJST(new Date(lastProcessedMessageDate))) {
        try {
          // 件名をログに記録
          Logger.log(`処理中のメール件名: ${message.getSubject()}`);
          let body = message.getBody();
          let subject = message.getSubject();
          let sender = message.getFrom();
          let senderEmail = sender.match(/<(.*?)>/)[1];
          let threadId = thread.getId();
          let messageId = message.getId();
          if(checkText(subject,body) == true){
            let payload = {
              "body": encodeURIComponent(body),
              "subject": encodeURIComponent(subject),
              "sender": senderEmail,
              "model": "gpt-3.5-turbo"
            };
            let options = {
              method: "post",
              contentType: "application/json",
              payload: JSON.stringify(payload)  // JSON形式に変換
            };
            let response = UrlFetchApp.fetch(url, options);
            let answer = JSON.parse(response.getContentText());
            Logger.log('[answer]'+answer);
            // 決算説明会情報
            ticker = ''//answer['証券コード']
            disclose_date = answer['決算発表日']
            disclose_time = answer['決算発表時間']
            briefing = answer['決算説明会開催有無']
            briefing_date = answer['決算説明会開催日']
            briefing_stime = answer['決算説明会開始時間']
            briefing_etime = answer['決算説明会終了時間']
            speaker_post = answer['決算説明会スピーカー役職']
            speaker_name = answer['決算説明会スピーカー氏名']
            briefing_title = answer['決算説明会イベント名']
            briefing_url = answer['決算説明会申し込み用URL']

            // 日本時間のタイムスタンプをフォーマット
            let timestamp = new Date();
            let formattedTimestamp = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");

            // スプレッドシートに追記
            sh.appendRow([subject, ticker, disclose_date, disclose_time, 
                     briefing, briefing_date, briefing_stime, briefing_etime, 
                     speaker_post, speaker_name, 
                     briefing_title, briefing_url,threadId,messageId,formattedTimestamp,body,userEmail]);
          }
          
          //スクリプトプロパティにスレッドIDとメッセージIDを保存
          newData.push({
            threadId: threadId,
            messageId: messageId
          });
          Logger.log(newData);

          message.markRead();
          if (!threadHasLabel(thread, processedLabel)) {
            thread.addLabel(processedLabel);
          }

          // 最終処理日時を更新
          scriptProperties.setProperty("lastProcessedMessageDate", message.getDate().toISOString());
        } catch (error) {
          console.error(`メール処理中にエラーが発生しました: ${error.message}`);
        }
      }
    }
  }

  // 既存データに新しいデータを結合（重複を避けるためにフィルタリング）
  const combinedData = [...data, ...newData].filter(
    (item, index, self) =>
      self.findIndex(i => i.threadId === item.threadId && i.messageId === item.messageId) === index
  );

  // スクリプトプロパティに保存
  scriptProperties.setProperty(propertyKey, JSON.stringify(combinedData));
  console.log(`本日保存したデータ件数: ${combinedData.length}`);

  console.log("処理完了");
}

// スレッドにラベルが付いているか確認
function threadHasLabel(thread, label) {
  const labels = thread.getLabels().map(l => l.getName());
  return labels.includes(label.getName());
}

// JSTに変換する関数
function convertToJST(date) {
  return new Date(date.getTime() + 9 * 60 * 60 * 1000); // UTCからJSTに変換
}

// タイムアウトチェック関数
function isTimeout(startTime) {
  return new Date().getTime() - startTime > 5.5 * 60 * 1000; // 実行時間が5.5分を超えるか
}

// 処理日時を保存
function saveLastProcessedDate(scriptProperties, messages) {
  if (messages && messages.length > 0) {
    const lastProcessed = messages[messages.length - 1].getDate().toISOString();
    scriptProperties.setProperty("lastProcessedMessageDate", lastProcessed);
    console.log(`処理中断時の最後の処理日時（JST）: ${convertToJST(new Date(lastProcessed)).toISOString()}`);
  } else {
    console.log("処理中断時に保存する対象のメールがありません。");
  }
}
