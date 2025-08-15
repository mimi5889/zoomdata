function updateEmailInfoFromCode() {//企業メールアドレス一覧からメールアドレスを取得し、スプシに転記する
  const scriptProperties = PropertiesService.getScriptProperties();
  const companyEmailSheetId = scriptProperties.getProperty('COMPANY_EMAIL_SHEET_ID');

  const zoomSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const emailSheet = SpreadsheetApp.openById(companyEmailSheetId).getSheets()[0];
  const emailList = emailSheet.getDataRange().getValues(); // 説明会支援先_担当者email
  const zoomData = zoomSheet.getDataRange().getValues();   // Zoomウェビナー一覧

  for (let i = 1; i < zoomData.length; i++) {
    if (zoomData[i][5] == 1) continue;
    const topic = zoomData[i][2]; // C列：トピック
    const matches = [];

    for (let j = 1; j < emailList.length; j++) {
      const code = emailList[j][0]; 
      const name = emailList[j][1]; 
      const email = emailList[j][2];
      if (topic.includes(code)) {
        matches.push([code, name, email]);
      }
    }

    if (matches.length === 1) {
      // 1件だけ一致 → HIJに転記
      zoomSheet.getRange(i + 1, 8, 1, 3).setValues([matches[0]]);
    } else if (matches.length > 1) {
      // 複数一致 → H列に「※」、HIは空欄
      zoomSheet.getRange(i + 1, 8, 1, 3).setValues([['※', '', '']]);
    } else {
      // 一致なし → すべて空欄
      zoomSheet.getRange(i + 1, 8, 1, 3).clearContent();
    }
  }

  Logger.log("メールアドレスの更新処理が完了しました");
}
