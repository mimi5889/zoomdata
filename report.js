function webinarReportsTrigger(){//10分おき（5分おき？）トリガー設定

  try{

    //メインコードが実行中か確認する
    //実行中の場合は即終了
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(1000)) { // 1秒以内に取れなければ優先ジョブ中
      Logger.log("優先ジョブ中のためスキップ");
      return;// 即終了
    }
    const sp = PropertiesService.getScriptProperties();
    const currentRow = sp.getProperty('i');

    if (currentRow && currentRow !== '2') {
      Logger.log(`優先ジョブ実行中（currentRow=${currentRow}）のためスキップ`);
      return; // 即終了
    }

    const now = new Date();
    const hour = now.getHours(); // 現在の「時」（0〜23）

    // 23時から7時 の場合は実行しない
    if (hour >= 23 || hour < 7) {
      Logger.log("この時間帯（23:00〜7:00）は実行しません。");
      return;// 即終了
    }

    generateWebinarReports();//事後データ作成
    console.log('-----処理終了------');

  }catch(e){
    console.log('-----レポート作成　エラーを検知しました。------');
    console.log('エラー内容：'+e.message);
    //エラー通知
  }

}

function generateWebinarReports() {//データ取得CSV作成メインプログラム
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const scriptProperties = PropertiesService.getScriptProperties();
  const max_acccountIndex = parseInt(scriptProperties.getProperty('MAX_ACCOUNT_INDEX') || '4');


  for(let accountIndex = 1 ; accountIndex <= max_acccountIndex ; accountIndex ++){
    const props = PropertiesService.getScriptProperties();
    const zoomId = props.getProperty('ZOOM_ID_' + accountIndex);
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetId = scriptProperties.getProperty('SHEET_ID');
    const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
    const flgSheet = sheet.getSheetByName('除外');
    const exclusionIds = flgSheet.getRange('B2:B')//メールの自動送信を除外する証券コード
    .getValues()
    .flat()
    .filter(word => word); // 空でないものだけ

    const matchedIndexes = data
      .map((row, i) => row[0] === zoomId ? i : -1)//zoomIDが一致する行だけ動かす
      .filter(i => i !== -1);

    matchedIndexes.forEach(i => {
      const webinarId = data[i][1];//B ウェビナーID
      const topic = data[i][2];//C トピック
      const startTimeStr = data[i][3];//E　終了予定時刻
      const endTimeStr = data[i][4];//E　終了予定時刻
      const endTimeReal = data[i][5];//E　終了時刻
      const exclusionFlag = data[i][6];//F 除外フラグ
      const stockId = data[i][7];//G　証券コード
      const companyName = data[i][8];//H　企業名
      const companyAdd = data[i][9];//I 企業送付先アドレス
      const attendeeLink = data[i][10];//J　出席者レポート
      const surveyLink = data[i][11];//I アンケート結果レポート
      const qaLink = data[i][12];//J　Q&Aレポート

      if (!webinarId || !endTimeStr) return;
      if (exclusionFlag !== '') return;
      if (attendeeLink || surveyLink || qaLink) return;

      Logger.log(i+':'+topic);

      const startTime = new Date(startTimeStr);  
      const endTime = new Date(endTimeStr);
      const diffMin = (now - endTime) / (1000 * 60);
      //Logger.log('diffMin:' + diffMin);

      //終了予定時刻45分前から120分以内でウェビナーが終了になるか監視する
      if (diffMin > -45 && diffMin < 120 && endTimeReal == "" && startTime < now){
        const ts = fetchWebinarReturnTime(webinarId,accountIndex);
        if (!ts){
          // データ取得できなかったらスキップ
          Logger.log(`skip: webinarId=${webinarId} index=${accountIndex}`);
        }else{
          const [y, mo, d, h, mi, s] = ts.match(/\d+/g).map(Number);
          const dateA = new Date(y, mo - 1, d, h, mi, s);   
          const dateB = new Date(startTime); 

          if (dateA.getTime() > dateB.getTime()) {
            sheet.getRange(i+1,6).setValue(ts);
          }
        }
      } 
      const realDiffMin = (now - endTimeReal) / (1000 * 60);  
      //終了時刻から35分以上60分以内で実行する
      if (realDiffMin < 35 || realDiffMin >60) return;
      const result = exportWebinarCsvs(webinarId, accountIndex, stockId, companyName, endTime,companyAdd);//*****ウェビナーデータからCSV作成*******
      sheet.getRange(i+1,11).setValue(result.fileUrls[0]);
      sheet.getRange(i+1,12).setValue(result.fileUrls[1]);
      sheet.getRange(i+1,13).setValue(result.fileUrls[2]);

      Logger.log('----exclusionIds----');
      Logger.log(exclusionIds);
      if(!exclusionIds.includes(stockId)){
        createDraftMail(stockId,companyName,companyAdd,result.attendeeFile,result.surveyFile,result.qaFile);//****下書きメール作成****
      }

    });
  }
}


function fetchWebinarReturnTime(webinarId,accountIndex) {//終了時刻判定
  const token = getAccessToken(accountIndex); // トークン取得関数に index を渡す
  if (!token) {
    Logger.log(`getAccessToken failed for index=${accountIndex}`);
    return null;
  }

  const url = `https://api.zoom.us/v2/metrics/webinars/${webinarId}`
             + '?type=past';

  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,    
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  const status = res.getResponseCode();        // HTTP ステータス
  const body   = res.getContentText();         // JSON 文字列
  Logger.log(`metrics/webinars status=${status}`);

  // --- 正常終了（200） ---
  if (status === 200) {
    const data = JSON.parse(body);
    Logger.log(data.topic);
    Logger.log(`終了時刻 : ${data.end_time}`);
    const jstNow = new Date(data.end_time);
    return Utilities.formatDate(jstNow, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  }else{
    return null;
  }
}



function exportWebinarCsvs(webinarId, accountIndex, stockId, companyName, endTime ,companyAdd) {//データを取得しcsv作成する
  try{
    const scriptProperties = PropertiesService.getScriptProperties();
    const folderId = scriptProperties.getProperty('FOLDER_ID');
    const token = getAccessToken(accountIndex); // トークン取得関数に index を渡す
    const dateStr = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'yyyyMMdd');
    const filePrefix = `${companyName}(${stockId})様`;

    // 1. トピックを取得
    const detail = fetchZoomData(`https://api.zoom.us/v2/webinars/${webinarId}`, token);
    const topic = detail.topic.replace(/[\\/:*?"<>|]/g, '_'); // フォルダ名に使えない文字を置換
    const qa = detail.settings.question_and_answer.enable;//Q&Aにチェックがあるかどうか
    Logger.log('qa' + qa);

    // 2. フォルダ作成
    const folder = getOrCreateFolderByName(folderId, topic);
    const folderUrl = folder.getUrl();

    // 3. 各CSV生成
    let webhooktxt = '';
    let attendeeFile = null;
    let surveyFile = null;
    let qaFile = null;

    //出席者レポート
    let csvAry = [];
    Logger.log('---------登録者レポート-----------');//登録者レポート
    const registrants = fetchZoomData(`https://api.zoom.us/v2/webinars/${webinarId}/registrants?page_size=1000`, token);
    Logger.log('----------出席者レポート----------');//参加者レポート
    const participantsData = fetchZoomData(`https://api.zoom.us/v2/report/webinars/${webinarId}/participants?page_size=1000`, token);
    Logger.log('----------パネリストレポート----------');//パネリストレポート
    const panelistsData = fetchZoomData(`https://api.zoom.us/v2/webinars/${webinarId}/panelists?page_size=1000`, token);
    Logger.log('---------ダッシュボード-----------');//ダッシュボード
    const dashBoadData = fetchZoomData(`https://api.zoom.us/v2/metrics/webinars/${webinarId}/participants` + '?type=past&page_size=300',token);

    const bom = '\uFEFF'; // UTF-8 BOMを付けることでExcelで誤認しにくくする

    if (registrants && registrants.registrants && registrants.registrants.length > 0) {//登録者データを探す
      const attendeeCsv = generateAttendeeCsv(registrants.registrants, participantsData?.participants || [] ,panelistsData?.panelists || [] ,accountIndex,dashBoadData?.participants || [],topic);//登録者データ、参加者データ、ダッシュボードデータを渡す
      const attendeeCsvWithBom = bom + attendeeCsv;
      const attendeeBlob = Utilities.newBlob(attendeeCsvWithBom, MimeType.CSV, `${filePrefix}-Attendee Report_${dateStr}.csv`);
      attendeeFile = folder.createFile(attendeeBlob);
      webhooktxt = webhooktxt + '\n'+'・出席者レポート';
      csvAry.push(attendeeFile.getUrl());
      Utilities.sleep(300);
    }else{
      csvAry.push('');//登録者データがない場合は空白で作成
    }

    // アンケート結果レポート
    Logger.log('----------アンケート結果レポート-----------');
    const result = validateZoomDataWithRetry(
      () => fetchZoomData(`https://api.zoom.us/v2/webinars/${webinarId}/survey`, token),
      (d) => d && d.custom_survey && Array.isArray(d.custom_survey.questions) && d.custom_survey.questions.length > 0
    );

    if(result.valid) { 
      const surveyDef = fetchZoomData(`https://api.zoom.us/v2/webinars/${webinarId}/survey`, token);//設問一覧
      Logger.log('---------設問一覧--------------');
      const customSurveyData = surveyDef.custom_survey;

      Logger.log('---------回答一覧--------------');
      const surveys = fetchZoomData(`https://api.zoom.us/v2/report/webinars/${webinarId}/survey`, token);//回答一覧
      const surveyAnswers = surveys.survey_answers;

      //アンケートの回答が無くても出力する
      const surveyCsv =  generateSurveyCsv(surveyAnswers,customSurveyData, topic);  // ← CSV生成関数も構造に合わせて
      const surveyCsvWithBom = bom + surveyCsv;
      const surveyBlob = Utilities.newBlob(surveyCsvWithBom, MimeType.CSV, `${filePrefix}-Survey Report_${dateStr}.csv`);
      surveyFile = folder.createFile(surveyBlob);
      webhooktxt += '\n・アンケート結果レポート';
      csvAry.push(surveyFile.getUrl());
      Utilities.sleep(300);
    }else{
      csvAry.push('');//空白で作成
    }


    if(qa == true){
      // Q&A結果レポート
      Logger.log('--------Q&A結果レポート---------');
      const qas = fetchZoomData(`https://api.zoom.us/v2/report/webinars/${webinarId}/qa`, token);
      const qaCsv = generateQaCsv(qas.questions,topic);
      const qaCsvWithBom = bom + qaCsv;
      const qaBlob = Utilities.newBlob(qaCsvWithBom, MimeType.CSV, `${filePrefix}-Q&A Report_${dateStr}.csv`);
      qaFile = folder.createFile(qaBlob);
      webhooktxt = webhooktxt + '\n'+'・Q&A結果レポート';
      csvAry.push(qaFile.getUrl());
      Utilities.sleep(300);
    }else{
      csvAry.push('');//空白で作成
    }

    Logger.log('csvAry:'+csvAry);
    
    sendSlackNotification(topic,folderUrl,webhooktxt,stockId,companyAdd);//***slack通知******************

    return {
      attendeeFile,
      surveyFile,
      qaFile,
      fileUrls: csvAry
    };

  }catch(e){
    Logger.log(e.message);
  }
}

function generateAttendeeCsv(registrants, reportParticipants, panelists, accountIndex, metricsParticipants,topic) {//出席者レポート
  const props = PropertiesService.getScriptProperties();
  const zoomId = props.getProperty('ZOOM_ID_' + accountIndex);
  const panelistEmails = new Set(panelists.map(p => (p.email || '').toLowerCase()));

  // 登録者・レポート参加者のマップ
  const registrantMap = new Map();
  registrants.forEach(r => {
    const email = (r.email || '').toLowerCase();
    if (email) registrantMap.set(email, r);
  });
  // カスタム質問タイトル
  const customTitlesSet = new Set();
  registrants.forEach(p => {
    (p.custom_questions || []).forEach(q => customTitlesSet.add(q.title));
  });
  const customTitles = Array.from(customTitlesSet);

  const headers = [
    '参加済み','トピック名', 'ユーザー名（オリジナル名）', '名（登録）', '姓（登録）', 'メール',
    '登録時間', '承認ステータス', '参加時間', '退出時間', 'セッション時間（分）',
    'は外部参加者',
    ...customTitles,
    '国/地域','会社名','役職','電話番号'
  ];

  const rows = [];

  metricsParticipants.forEach(mp => {
    const email = (mp.email || '').toLowerCase();
    //Logger.log(email);
    const name = mp.user_name || '';

    if (panelistEmails.has(email)) {
      // パネリストに登録されているメールはスキップ
      return;
    }
    if (email === zoomId.toLowerCase()){
      return;
    }
    const reg = registrantMap.get(email) || {};
    const participated = 'はい';
    const isGuest = mp.role === 'attendee' ? 'はい' : '-';
    const join_time = mp.join_time ? Utilities.formatDate(new Date(mp.join_time), 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss') : '-';
    const leave_time = mp.leave_time ? Utilities.formatDate(new Date(mp.leave_time), 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss') : '-';

    let durationMin = '';
    if (mp.join_time && mp.leave_time) {
      const join = new Date(mp.join_time);
      const leave = new Date(mp.leave_time);
      const diffMs = leave - join;
      if (!isNaN(diffMs)) {
        durationMin = Math.ceil(diffMs / 60000); // 1000ms * 60秒 = 1分
      }
    }
    const create_time = reg.create_time ? Utilities.formatDate(new Date(reg.create_time), 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss') : '-';
    const status = reg.status === 'approved' ? '承認済み' : reg.status || '';
    const location = mp.location || '';

    const customAnswersMap = {};
    (reg.custom_questions || []).forEach(q => {
      customAnswersMap[q.title] = q.value;
    });
    const customAnswers = customTitles.map(title => customAnswersMap[title] || '');

    rows.push([
      participated,
      topic,
      name || '',
      reg.first_name || '',
      reg.last_name || '',
      email || '',
      create_time,
      status,
      join_time,
      leave_time,
      durationMin,
      isGuest,
      ...customAnswers,
      location,
      reg.org || '',
      reg.job_title || '',
      reg.phone || ''
    ]);
  });

  // metricsParticipants に含まれる user_id をセットに保持
  const existingUserIds = new Set(metricsParticipants.map(mp => (mp.user_id || '').toLowerCase()));

  // rows からも user_id を取得（metricsで取得したデータとrowsに追加したデータの両方）
  rows.forEach(row => {
    const userId = (row.__user_id || '').toLowerCase(); // 後述：__user_idをrowに一時保持しておく場合
    if (userId) existingUserIds.add(userId);
  });

  reportParticipants.forEach(rp => {
    const userId = (rp.user_id || '').toLowerCase();
    if (!userId || existingUserIds.has(userId)) return;

    const participated = 'はい';
    const isGuest = rp.role === 'attendee' ? 'はい' : '-';
    const join_time = rp.join_time ? Utilities.formatDate(new Date(rp.join_time), 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss') : '-';
    const leave_time = rp.leave_time ? Utilities.formatDate(new Date(rp.leave_time), 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss') : '-';

    let durationMin = '';
    let location = "";
    if (rp.join_time && rp.leave_time) {
      const join = new Date(rp.join_time);
      const leave = new Date(rp.leave_time);
      const diffMs = leave - join;
      if (!isNaN(diffMs)) {
        durationMin = Math.ceil(diffMs / 60000);
      } 
    }
    const isPhoneOnly = !rp.email && /^\d{8,}$/.test(rp.name);
    if (isPhoneOnly) location = inferCountryISO(rp.name);
    rows.push([
      'はい',
      topic,
      rp.name || '',
      '', '', // 名、姓
      rp.email || '',
      '-', // 登録時間
      '', // 承認ステータス
      join_time,
      leave_time,
      durationMin,
      'はい',
      ...customTitles.map(() => ''), // カスタム質問なし
      location,
      rp.org || '',
      rp.job_title || '',
      rp.phone || ''
    ]);

    existingUserIds.add(userId);
  });


  rows.sort((a, b) => {
    // 1列目「参加済み」
    const primary = b[0].localeCompare(a[0]); // 'はい' > 'いいえ'
    if (primary !== 0) return primary;

    // 3列目「ユーザー名（オリジナル名）」
    const secondary = a[2].localeCompare(b[2]);
    if (secondary !== 0) return secondary;

    // 8列目「参加時間」
    const dateA = new Date(a[8]);
    const dateB = new Date(b[8]);
    return dateA - dateB;
  });

  // 参加していない登録者を追加（参加者リストに含まれていないregistrants）
  const participantEmails = new Set(metricsParticipants.map(p => p.email));

  registrants.forEach(r => {
    if (!participantEmails.has(r.email) && !panelistEmails.has(r.email) && r.email !== zoomId) {
      const create_time = r.create_time
        ? Utilities.formatDate(new Date(r.create_time), 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss')
        : '-';
      const status = r.status === 'approved' ? '承認済み' : r.status || '';
      const customAnswersMap = {};
      (r.custom_questions || []).forEach(q => {
        customAnswersMap[q.title] = q.value;
      });
      const customAnswers = customTitles.map(title => customAnswersMap[title] || '');

      rows.push([
        'いいえ',              // 参加していない
        topic,
        '',                    // ユーザー名（不明）
        r.first_name || '',
        r.last_name || '',
        r.email || '',
        create_time,
        status,
        '-', '-', '',          // 参加・退出・duration
        '-',                   // isGuest
        ...customAnswers,
        '' ,'','' ,''                  // location
      ]);
    }
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

    // 降順にソートして後ろから削除（インデックスずれを防ぐため）
    emptyColIndexes.sort((a, b) => b - a).forEach(index => {
      headers.splice(index, 1);
      rows.forEach(row => row.splice(index, 1));
    });
  }

  return [headers, ...rows]
    .map(row => row.map(cell => `"${sanitizeCell(cell)}"`).join(','))
    .join('\n');
}


function generateSurveyCsv(surveyAnswers, customSurveyData, topic) {//アンケートレポート
  const questionList = [];

  // 設問リストを name 順に登録（重複許容）
  if (customSurveyData && Array.isArray(customSurveyData.questions)) {
    customSurveyData.questions.forEach(q => {
      if (q && q.name) questionList.push(q.name);
    });
  }

  const headers = ['メール', '名前', '回答日時', 'トピック名', ...questionList];

  if (!Array.isArray(surveyAnswers) || surveyAnswers.length === 0) {
    return [headers]
      .map(row => row.map(cell => `"${sanitizeCell(cell)}"`).join(','))
      .join('\n');
  }

  const rows = surveyAnswers.map(answer => {
    const email = answer.email || '';
    const name = answer.name || '';
    const dateObj = answer.date_time ? new Date(answer.date_time) : null;
    const sortKey = dateObj && !isNaN(dateObj.getTime()) ? dateObj.getTime() : 0;
    const dateTime = dateObj
      ? Utilities.formatDate(dateObj, 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss')
      : '-';

    const answerDetails = Array.isArray(answer.answer_details) ? answer.answer_details : [];
    const usedIndices = new Set(); // 重複マッチ防止

    const orderedAnswers = questionList.map(qName => {
      for (let i = 0; i < answerDetails.length; i++) {
        const detail = answerDetails[i];
        if (
          detail &&
          detail.question === qName &&
          !usedIndices.has(i)
        ) {
          usedIndices.add(i);
          return detail.answer || '';
        }
      }
      return ''; // 該当なし
    });

    const answerRow = [email, name, dateTime, topic, ...orderedAnswers];
    return { sortKey, answerRow };
  });

  rows.sort((a, b) => a.sortKey - b.sortKey);

  const csv = [headers, ...rows.map(r => r.answerRow)]
    .map(row => row.map(cell => `"${sanitizeCell(cell)}"`).join(','))
    .join('\n');

  return csv;
}


function generateQaCsv(questions,topic) {//Q&A結果レポート
  const headers = [
    'トピック名',
    '質問',
    '質問者名',
    '質問者のメール',
    '回答',
    '質問時間',
    '応答した時間',
    '回答名',
    'メールに応答'
  ];

  const rowObjs = [];

  if (Array.isArray(questions) && questions.length > 0) {
    questions.forEach(p => {
      const name = p.name || '';
      const email = p.email || '';

      if (Array.isArray(p.question_details)) {
        p.question_details.forEach(qd => {
          const question = qd.question || '';
          const questionDateObj = qd.create_time ? new Date(qd.create_time) : null;
          const questionTime = questionDateObj
            ? Utilities.formatDate(questionDateObj, 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss')
            : '-';

          if (Array.isArray(qd.answer_details) && qd.answer_details.length > 0) {
            qd.answer_details.forEach(ad => {
              const answerTime = ad.create_time
                ? Utilities.formatDate(new Date(ad.create_time), 'Asia/Tokyo', 'EEE MMM dd yyyy HH:mm:ss')
                : '';
              rowObjs.push({
                sortKey: questionDateObj ? questionDateObj.getTime() : 0,
                row: [
                  topic,
                  question,
                  name,
                  email,
                  ad.content || '',
                  questionTime,
                  answerTime,
                  ad.name || '',
                  ad.email || ''
                ]
              });
            });
          } else {
            rowObjs.push({
              sortKey: questionDateObj ? questionDateObj.getTime() : 0,
              row: [
                topic,
                question,
                name,
                email,
                '',
                questionTime,
                '',
                '',
                ''
              ]
            });
          }
        });
      }
    });
  }

  // 質問時間で昇順ソート
  rowObjs.sort((a, b) => a.sortKey - b.sortKey);

  // CSV文字列へ変換
  const csv = [headers, ...rowObjs.map(r => r.row)]
    .map(row => row.map(cell => `"${sanitizeCell(cell)}"`).join(','))
    .join('\n');

  return csv;
}


function fetchZoomData(url, token) {//ZooAPI接続

  const scriptProperties = PropertiesService.getScriptProperties();
  const logSheetId = scriptProperties.getProperty('LOG_SHEET_ID');
  const logSs = SpreadsheetApp.openById(logSheetId);
  const logSh = logSs.getSheets()[0]; // 一番左のシート

  try {
    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true  // 👈 これで404などのレスポンスも受け取れる
    });

    const code = response.getResponseCode();
    const body = response.getContentText();

    Logger.log(`📡 URL: ${url}`);
    Logger.log(`🔐 Token (start): ${token.substring(0, 10)}...`);
    Logger.log(`📥 Status: ${code}`);
    Logger.log(`📄 Body: ${body}`);

    const logSh_lastRow = logSh.getLastRow();
    const maxLength = 50000;

    const chunks = [];

    for (let i = 0; i < body.length; i += maxLength) {
      chunks.push(body.substring(i, i + maxLength));
    }
    const values2D = chunks.map(c => [c]); // 各chunkを1列の2次元配列に変換

    //logSh.getRange(logSh_lastRow+1,1).setValue(new Date());
    //logSh.getRange(logSh_lastRow+1,2).setValue(url);
    //logSh.getRange(logSh_lastRow + 1, 3, chunks.length, 1).setValues(values2D);

    if (code === 200) {
      return JSON.parse(body);
    } else {
      Logger.log(`❌ Zoom APIエラー - ステータスコード: ${code}`);
      return null;
    }
  } catch (e) {
    Logger.log(`⚠️ fetchZoomData エラー: ${e}`);
    return null;
  }
}

function getOrCreateFolderByName(parentFolderId, topic) {//フォルダ作成
  const parent = DriveApp.getFolderById(parentFolderId);
  const folders = parent.getFoldersByName(topic);

  let folder;
  if (folders.hasNext()) {
    folder = folders.next(); // 既存のフォルダを使用
    Logger.log(`既存フォルダを使用: ${folder.getName()}`);
  } else {
    folder = parent.createFolder(topic); // 新規作成
    Logger.log(`新しいフォルダを作成: ${folder.getName()}`);
  }

  return folder; // 必要に応じて folder.getUrl() などにしてもOK
}

function sanitizeCell(cell) {//改行を削除
  return typeof cell === 'string'
    ? cell.replace(/[\r\n]+/g, ' ') // 改行をスペースに置換（または空文字でも可）
    : cell;
}

/**
 * 設問の有無をバリデーションして、取得に失敗した場合はリトライする
 *
 * @param {function(): Object} fetchFn - データ取得関数（例: () => fetchZoomData(url, token)）
 * @param {function(Object): boolean} validateFn - バリデーション関数（true なら有効）
 * @param {number} [retryMax=3] - 最大リトライ回数
 * @return {Object} { valid: boolean, data: any }
 */
function validateZoomDataWithRetry(fetchFn, validateFn, retryMax = 3) {
  for (let attempt = 1; attempt <= retryMax; attempt++) {
    try {
      const data = fetchFn();
      if (validateFn(data)) {
        Logger.log(`データ取得成功（${attempt}回目）`);
        return { valid: true, data };
      } else {
        Logger.log(`データ形式不正（${attempt}回目）`);
      }
    } catch (e) {
      Logger.log(`データ取得エラー（${attempt}回目）: ${e}`);
    }
    Utilities.sleep(1000); // 少し待つ
  }
  return { valid: false, data: null };
}

/**
 * 電話番号から ISO-2 国コードを推定
 * - + が無い／国際発信プレフィックスが欠落している番号を考慮
 * @param {string} raw  e.g. '818012345678', '11818012345678', '0312345678'
 * @return {string} ISO-2 (大文字) or '' (= 推定不能)
 */
function inferCountryISO(raw) {
  if (!raw) return '';
  
  // 1) 数字だけ残す
  let num = raw.replace(/\D/g, ''); 
  
  // 2) 国際発信プレフィックス除去
  if (num.startsWith('00'))  num = num.slice(2);   // 00xxxx…
  else if (num.startsWith('011')) num = num.slice(3); // 011xxxx…
  else if (num.startsWith('0011')) num = num.slice(4); // 0011xxxx…
  else if (num.startsWith('11') && num.length > 11) num = num.slice(2); // 11 + 国コード…
  // (+ が無い E.164 はそのまま)

  // 3) 0 で始まるなら国内表記 → 国判定不可
  if (num.startsWith('0')) return '';

  // 4) 1〜3桁で最長一致
  for (let len = 3; len >= 1; len--) {
    const code = num.slice(0, len);
    if (COUNTRY_CODE_MAP[code]) return COUNTRY_CODE_MAP[code].iso;
  }
  return '';
}
// 主要国番号（抜粋）— 必要に応じて拡張
const COUNTRY_CODE_MAP = {
  '1':  { iso: 'US' },   // NANP
  '44': { iso: 'GB' },
  '61': { iso: 'AU' },
  '65': { iso: 'SG' },
  '81': { iso: 'JP' },
  '86': { iso: 'CN' },
  '49': { iso: 'DE' },
  '33': { iso: 'FR' },
  '39': { iso: 'IT' },
  '34': { iso: 'ES' },
  // 全リストは GitHub などの JSON を都度読み込むと保守が楽
};




