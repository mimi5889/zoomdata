# ウェビナー自動データ抽出システム

## 概要
Google Apps Scriptを使用したZoomウェビナーの事前・事後データ自動取得システム

## セキュリティ修正履歴

### 2025年8月14日 - セキュリティ強化
- **機密情報の除去**: ハードコードされたID、メールアドレス、APIキーをすべて除去
- **スクリプトプロパティ化**: 設定値をスクリプトプロパティに移動
- **セキュリティ向上**: コードの再利用性と保守性が大幅に向上

#### 修正された項目
- スプレッドシートID
- フォルダID
- メールアドレス
- Slack Webhook URL
- Zoom API設定
- アカウント数設定

#### 新しく必要なスクリプトプロパティ
- `SHEET_ID`
- `COMPANY_EMAIL_SHEET_ID`
- `FOLDER_ID`
- `CC_EMAIL`
- `SLACK_WEBHOOK_URL`
- `MAX_ACCOUNT_INDEX`
- `LOG_SHEET_ID`
- `ZOOM_ID_1` 〜 `ZOOM_ID_4`
- `CLIENT_ID_1` 〜 `CLIENT_ID_4`
- `CLIENT_SECRET_1` 〜 `CLIENT_SECRET_4`
- `ACCOUNT_ID_1` 〜 `ACCOUNT_ID_4`

### 2025年8月15日 - コンテナバインドスクリプト最適化
- **SHEET_ID参照の削除**: メインスプレッドシートの参照を`getActiveSpreadsheet()`に変更
- **コンテナバインド対応**: スプレッドシートに紐づいたスクリプトとして最適化
- **実行環境の統一**: すべての関数でアクティブスプレッドシートを直接参照

#### 変更されたファイル
- `registants.js` - 17行目
- `report.js` - 41行目、42行目
- `test.js` - 11行目、15行目、105行目、110行目
- `webhook.js` - 3行目、4行目
- `webinarList.js` - 31行目、100行目、153行目
- `companyAdd.js` - 4行目

#### 変更内容
```javascript
// 変更前
const sheetId = scriptProperties.getProperty('SHEET_ID');
const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];

// 変更後
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
```

#### 残っているスクリプトプロパティ（必要）
- `COMPANY_EMAIL_SHEET_ID` - 別のシート用
- `FOLDER_ID` - Google Driveフォルダ用
- `MAX_ACCOUNT_INDEX` - Zoomアカウント設定用
- `SLACK_WEBHOOK_URL` - Slack通知用
- `CC_EMAIL` - メールCC用
- `ZOOM_ID_1` 〜 `ZOOM_ID_4` - Zoom API認証用
- `CLIENT_ID_1` 〜 `CLIENT_ID_4` - Zoom API認証用
- `CLIENT_SECRET_1` 〜 `CLIENT_SECRET_4` - Zoom API認証用
- `ACCOUNT_ID_1` 〜 `ACCOUNT_ID_4` - Zoom API認証用

### 2025年8月15日 - Googleサーバーエラー対応と軽量テスト機能追加
- **Googleサーバーエラー対応**: 一時的なサーバーエラー時の自動再実行機能を追加
- **進捗管理強化**: 処理中断時の進捗保存と再開機能
- **軽量テスト機能**: 副作用のない安全なテスト環境を実装
- **エラーハンドリング改善**: 個別行エラーと全体エラーの適切な分離

#### 新機能の詳細

##### Googleサーバーエラー対応
- **対象エラー**: Service temporarily unavailable, Internal error, Quota exceeded, Rate Limit Exceeded
- **再実行制御**: 最大3回まで自動再実行（1分間隔）
- **進捗保存**: 中断した行から継続処理
- **Slack通知**: エラー詳細と再実行状況の通知

##### 軽量テスト機能
- **事前データ取得テスト（軽量）**: メール送信・CSV作成・Driveアップロードなし
- **事後データ取得テスト（軽量）**: CSV作成・Driveアップロード・メール下書き作成なし
- **Webhookテスト（軽量）**: 実際の業務通知なし（テスト用軽量通知のみ）
- **安全性**: スプレッドシートへの記入・更新なし

#### 変更されたファイル
- `registants.js` - Googleサーバーエラー対応、進捗管理、再実行機能
- `test.js` - 軽量テスト関数3種類、メニュー更新

#### 実装された機能
```javascript
// 進捗管理
scriptProperties.setProperty('i', String(parseInt(i + 1, 10)));
scriptProperties.setProperty('slackAry', JSON.stringify(slackAry));

// Googleサーバーエラー検出と再実行
if (rowError.message.includes('Service temporarily unavailable') || 
    rowError.message.includes('Internal error') ||
    rowError.message.includes('Quota exceeded') ||
    rowError.message.includes('Rate Limit Exceeded')) {
  
  if (errorCount >= 3) {
    // 3回以上エラーが発生した場合は中止
    webhooktxt = `⚠️Googleサーバーエラーが3回発生しました\n処理を中止します\n現在の行: ${i}\nエラー: ${rowError.message}`;
    sendSlackNotification2(webhooktxt);
    scriptProperties.setProperty('i', '2');
    scriptProperties.setProperty('slackAry', 'NaN');
    scriptProperties.setProperty('errorCount', '0');
    return;
  } else {
    // 再実行を試行
    webhooktxt = `⚠️Googleサーバーエラーが発生しました\n1分後に再実行します\n現在の行: ${i}\nエラー: ${rowError.message}\n再実行回数: ${errorCount + 1}/3`;
    sendSlackNotification2(webhooktxt);
    scriptProperties.setProperty('errorCount', String(errorCount + 1));
    setRetryTrigger();
    return;
  }
}
```

#### 軽量テストの利点
- **完全に安全**: データを一切変更しない
- **高速実行**: 記入処理がないため高速
- **繰り返し実行可能**: 何度でも実行できる
- **本番環境でも安心**: データに影響しない
- **設定確認**: スクリプトプロパティ、API接続、データ取得の動作確認

## セットアップ手順
1. スクリプトプロパティに必要な値を設定
2. 軽量テストで基本動作を確認
3. 各機能のテスト実行
4. 本格運用開始

## 注意事項
- 機密情報はすべてスクリプトプロパティで管理
- コードに直接機密情報を記載しない
- 環境別設定が可能
- **コンテナバインドスクリプト**: スプレッドシートを開いた状態で実行する必要があります
- **SHEET_ID不要**: メインスプレッドシートは自動的にアクティブシートとして認識されます
- **軽量テスト推奨**: 本番環境での動作確認前に軽量テストを実行してください
- **エラー対応**: Googleサーバーエラー時は自動的に再実行されます（最大3回）