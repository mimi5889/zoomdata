function getAccessToken(index) {//インデックスごとにトークンを取得
  Logger.log('index:'+index);
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('CLIENT_ID_' + index);
  const clientSecret = props.getProperty('CLIENT_SECRET_' + index);
  const accountId = props.getProperty('ACCOUNT_ID_' + index);
  const tokenUrl = `https://zoom.us/oauth/token?grant_type=account_credentials&account_id=${accountId}`;
  const res = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret)
    }
  });
  const data = JSON.parse(res.getContentText());
  return data.access_token;
}

function getZoomUserId(token) {//トークン取得後にuserIDを取得する
  const response = UrlFetchApp.fetch('https://api.zoom.us/v2/users/me', {
    headers: {
      Authorization: 'Bearer ' + token
    }
  });
  return JSON.parse(response.getContentText()).id;
}
