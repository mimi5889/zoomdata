function createRegistantsMail(stockId, companyName, companyAdd, file) {
  const attachments = [];
  Logger.log('companyAdd:' + companyAdd);
  const to = companyAdd;
  const scriptProperties = PropertiesService.getScriptProperties();
  const cc = scriptProperties.getProperty('CC_EMAIL');
 

  const subject = '＜みんせつ＞【自動送信】事前データ送付のご案内';
  let body = '【' + companyName + '(' + stockId + ')】\n'
          + 'ご担当者さま\n\n'
          + '参加者リストを添付ファイルにてお送りいたします。\n'
          + 'ご不明点などございましたら、担当までお問合せください。';

  attachments.push(file.getBlob());
  
  GmailApp.sendEmail(to, subject, body, {
    attachments: attachments,
    cc: cc,
    name: '株式会社みんせつ'
  });
  Utilities.sleep(5000);
}

function createDraftMail(stockId, companyName, companyAdd, attendeeFile, surveyFile, qaFile) {
  const attachments = [];
  Logger.log('companyAdd:' + companyAdd);
  const to = companyAdd;
  const scriptProperties = PropertiesService.getScriptProperties();
  const cc = scriptProperties.getProperty('CC_EMAIL');
  const subject = '＜みんせつ＞【自動送信】事後データ送付のご案内';
  let body = '【' + companyName + '(' + stockId + ')】\n'
          + 'ご担当者さま\n\n'
          + '以下のデータを添付ファイルにてお送りいたします。\n';

  if (attendeeFile != null) {
    body += '・出席者リスト\n';
    attachments.push(attendeeFile.getBlob());
  }

  if (qaFile != null) {
    body += '・Q&A\n';
    attachments.push(qaFile.getBlob());
  }

  if (surveyFile != null) {
    body += '・アンケート\n';
    attachments.push(surveyFile.getBlob());
  }

  body += '\nご不明点などございましたら、担当までお問合せください。';
  //createDraft
  GmailApp.sendEmail(to, subject, body, {
    attachments: attachments,
    cc: cc,
    name: '株式会社みんせつ'
  });
}


