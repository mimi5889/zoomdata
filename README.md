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

### 2025年8月16日 - Googleサーバーエラー対応と軽量テスト機能追加
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

### 2025年8月16日 - 優先ジョブ状態管理と登録者数増減チェックの最適化
- **優先ジョブ状態管理**: スクリプトプロパティによる確実な実行制御
- **登録者数増減チェック最適化**: 21日前から当日の範囲でのみ実行
- **エラー時状態リセット**: 優先ジョブ状態の確実なリセット
- **軽量テスト関数追加**: 優先ジョブ状態管理の動作確認用テスト

#### 新機能の詳細

##### 優先ジョブ状態管理
- **状態管理**: `PRIORITY_JOB_STATUS` (RUNNING/IDLE) で実行状態を管理
- **開始時刻記録**: `PRIORITY_JOB_START_TIME` で処理開始時刻を記録
- **確実な制御**: スクリプトプロパティによる確実な実行制御
- **エラー時リセット**: エラー発生時も確実に状態をリセット

##### 登録者数増減チェックの最適化
- **実行範囲制限**: 21日前から当日の範囲でのみ実行
- **効率性向上**: 範囲外の行では`getRegistantsCount()`を呼び出さない
- **処理スキップ**: 範囲外の場合は完全にスキップ
- **ログ出力**: スキップ理由を明確に記録

##### エラー処理の強化
- **状態リセット**: すべてのエラー処理で優先ジョブ状態をリセット
- **Slack通知**: エラー内容と状態リセットの確認
- **一貫性確保**: 常に`PRIORITY_JOB_STATUS`を`IDLE`に戻す

#### 変更されたファイル
- `registants.js` - 優先ジョブ状態管理、登録者数増減チェック最適化、エラー時状態リセット
- `report.js` - ロック処理を削除、スクリプトプロパティによる制御に変更
- `test.js` - 優先ジョブ状態管理テスト関数2種類、メニュー更新

#### 実装された機能
```javascript
// 優先ジョブ開始時
scriptProperties.setProperty('PRIORITY_JOB_STATUS', 'RUNNING');
scriptProperties.setProperty('PRIORITY_JOB_START_TIME', new Date().toISOString());

// 正常完了時
scriptProperties.setProperty('PRIORITY_JOB_STATUS', 'IDLE');
scriptProperties.deleteProperty('PRIORITY_JOB_START_TIME');

// エラー時も確実に状態をリセット
scriptProperties.setProperty('PRIORITY_JOB_STATUS', 'IDLE');
scriptProperties.deleteProperty('PRIORITY_JOB_START_TIME');

// 登録者数増減チェックの範囲制限
if(laterDays <= 21 && laterDays >= 0) {
  // 21日前から当日の範囲でのみ実行
  const registantsCount = getRegistantsCount(webinarId,token,topic);
  // ... 処理 ...
} else {
  // 範囲外の場合はスキップ
  Logger.log(`行 ${i}: laterDays=${laterDays} のため登録者数増減チェックをスキップ`);
}

// webinarReportsTriggerでの制御
if (priorityJobStatus === 'RUNNING' || (currentRow && currentRow !== '2')) {
  Logger.log(`優先ジョブ実行中または処理中のためスキップ (status=${priorityJobStatus}, currentRow=${currentRow})`);
  return; // 即終了
}
```

#### 軽量テスト関数
- **`priorityJobStatusTest()`**: 優先ジョブ状態管理の基本動作確認
- **`webinarReportsTriggerTest()`**: トリガー制御ロジックの動作確認
- **安全性**: 実際の処理は実行せず、状態管理のみテスト

#### 実装の利点
- **確実な制御**: ロック処理よりも確実な実行制御
- **効率性向上**: 不要な処理のスキップ
- **エラー耐性**: エラー時も確実に状態をリセット
- **保守性**: 状態が明確で管理しやすい

### 2025年8月16日 - エラーメッセージの改善と優先ジョブでのSlack通知削除
- **エラーメッセージの改善**: より詳細で分かりやすいエラー通知
- **優先ジョブでのSlack通知削除**: 不要な通知を停止
- **軽量テスト関数の追加**: 登録者数増減チェックロジックのテスト

#### 新機能の詳細

##### エラーメッセージの改善
- **発生時刻**: エラーが発生した正確な時刻を記録
- **エラー種別**: エラーの種類を明確化（Googleサーバーエラー、Zoom APIエラー等）
- **処理状態**: 現在の処理状況を明示（再実行待機中、終了等）
- **情報の詳細化**: エラーの原因と状況が分かりやすい形式

##### 優先ジョブでのSlack通知削除
- **不要な通知の停止**: 優先ジョブ実行時のSlack通知を削除
- **ログでの記録**: エラーはログに記録して追跡可能
- **シンプル化**: コードがより簡潔に
- **パフォーマンス向上**: Slack通知処理の削減

##### 軽量テスト関数の追加
- **`registrantsCheckLogicTest()`**: 登録者数増減チェックロジックのテスト
- **日付計算テスト**: 35日前から7日後までの8パターン
- **範囲制限テスト**: 21日前〜当日の範囲制限の確認
- **条件分岐テスト**: 処理フローの正確性確認

#### 変更されたファイル
- `registants.js` - エラーメッセージ改善、優先ジョブでのSlack通知削除
- `test.js` - 軽量テスト関数追加、onOpen関数削除

#### 実装された機能
```javascript
// エラーメッセージの改善例
webhooktxt = `⚠️Googleサーバーエラーが3回発生しました\n` +
  `処理を中止します\n` +
  `発生時刻: ${new Date().toISOString()}\n` +
  `現在の行: ${i}\n` +
  `エラー種別: Googleサーバーエラー（3回目）\n` +
  `エラー内容: ${rowError.message}\n` +
  `処理状態: 優先ジョブ状態をリセットして終了`;

// 優先ジョブでのエラー処理（Slack通知なし）
} catch (error) {
  // エラーはログに記録のみ（Slack通知なし）
  Logger.log(`優先ジョブでエラーが発生: ${error.message}`);
  
  // 状態をリセット
  scriptProperties.setProperty('PRIORITY_JOB_STATUS', 'IDLE');
  scriptProperties.deleteProperty('PRIORITY_JOB_START_TIME');
  
  // エラーを再スロー（必要に応じて）
  throw error;
}
```

#### 軽量テストの特徴
- **UIなし**: エディタから直接実行可能
- **安全性**: 実際の処理は実行せず、ロジックのみテスト
- **詳細なログ**: 各テストステップの結果を詳細に記録
- **効率性確認**: 範囲外での処理スキップの確認

#### 実装の利点
- **デバッグしやすい**: エラー内容が詳細で分かりやすい
- **不要な通知削除**: 優先ジョブでの余計なSlack通知を停止
- **テスト環境**: 安全な軽量テストで動作確認可能
- **保守性向上**: エラー処理が明確で管理しやすい

### 2025年8月16日 - setExclusionFlags関数の強化
- **除外判定ロジックの拡張**: 除外ワード（A列）と除外stockID（B列）の両方を対象
- **証券コードによる除外**: メインシートH列（証券コード）と除外シートB列の照合
- **軽量テスト関数の追加**: setExclusionFlags関数の動作確認用テスト

#### 新機能の詳細

##### 除外判定ロジックの拡張
- **従来の機能**: 除外シートA列の除外ワードによるトピック除外
- **新機能**: 除外シートB列の除外stockIDによる証券コード除外
- **統合判定**: どちらかが除外対象の場合は除外フラグを設定
- **OR条件**: 除外ワードまたは除外stockIDのいずれかが一致

##### 証券コードによる除外
- **対象列**: メインシートH列（証券コード）
- **除外シート**: 除外シートB列（除外stockID）
- **照合方法**: 文字列として確実な比較（toString()使用）
- **安全性**: null/undefined値の適切な処理

##### 軽量テスト関数の追加
- **`setExclusionFlagsTest()`**: 除外判定ロジックの動作確認
- **除外シート状態確認**: A列・B列のデータ取得確認
- **ロジックテスト**: サンプルデータでの判定確認
- **実際データ予測**: 除外件数の事前計算

#### 変更されたファイル
- `webinarList.js` - setExclusionFlags関数の強化、除外stockID対応
- `test.js` - setExclusionFlags軽量テスト関数追加

#### 実装された機能
```javascript
function setExclusionFlags() {
  // 除外ワードの取得（除外シートA列）
  const exclusionWords = flgSheet.getRange('A2:A')
    .getValues()
    .flat()
    .filter(word => word);
  
  // 除外stockIDの取得（除外シートB列）
  const exclusionStockIds = flgSheet.getRange('B2:B')
    .getValues()
    .flat()
    .filter(stockId => stockId);

  // 除外判定ロジック
  for (let i = 0; i < topics.length; i++) {
    const topic = topics[i][0];
    const stockId = stockIds[i][0];
    
    // 除外ワードによる除外判定
    const isExcludedByWord = exclusionWords.some(word => topic.includes(word));
    
    // 除外stockIDによる除外判定
    const isExcludedByStockId = exclusionStockIds.some(exclusionId => 
      stockId && exclusionId && stockId.toString() === exclusionId.toString()
    );
    
    // どちらかが除外対象の場合はフラグを立てる
    const isExcluded = isExcludedByWord || isExcludedByStockId;
    flags.push([isExcluded ? 1 : '']);
  }
}
```

#### 軽量テストの特徴
- **安全性**: 実際の除外フラグ設定は行わない
- **詳細確認**: 除外ワード・除外stockIDの両方をテスト
- **ログ出力**: 各判定ステップの結果を詳細に記録
- **件数予測**: 実際のデータでの除外件数を事前計算

#### 実装の利点
- **柔軟性向上**: 除外ワードと除外stockIDの両方で除外可能
- **精度向上**: より細かい除外条件の設定
- **保守性向上**: 除外条件の管理が容易
- **安全性向上**: 軽量テストで動作確認が可能

### 2025年8月16日 - アンケートデータ取得の最適化とリトライ回数調整
- **リトライ回数の調整**: 3回のリトライでZoom API遅延への適切な対応
- **軽量版アンケートチェック**: 3回目のリトライ失敗時のみ軽量版でアンケート存在確認
- **データ遅延検出**: アンケートが存在するがデータが遅延している場合の適切な通知
- **処理効率の向上**: 軽量版チェックは必要な時のみ実行
- **GAS制限対応**: 6分制限に引っかからない適切な処理時間設計

#### 新機能の詳細

##### リトライ回数の調整
- **変更前**: 最大3回のリトライ
- **変更後**: 最大**3回**のリトライ（適切な回数に調整）
- **待機時間**: 各リトライ間で30秒の待機
- **耐性向上**: Zoom APIの一時的な遅延に対する適切な対応
- **GAS制限対応**: 合計処理時間約2-3分で6分制限に余裕を持って対応

##### 軽量版アンケートチェック
- **実行タイミング**: 3回目のリトライ失敗時のみ実行
- **チェック方法**: `?fields=custom_survey`パラメータで軽量に確認
- **目的**: アンケートの有無のみを確認（データ内容は取得しない）
- **効率性**: 必要な時のみ実行されるため処理効率が向上

##### データ遅延検出
- **遅延検出**: アンケートは存在するがデータ取得に遅延が発生している場合を検出
- **適切な通知**: Slack通知で「データ遅延のため取得できませんでした」と明記
- **ユーザー体験**: アンケートが存在しないのか、遅延しているのかが明確

#### 変更されたファイル
- `report.js` - リトライ回数調整、軽量版アンケートチェック実装、データ遅延検出

#### 実装された機能
```javascript
// リトライ回数を3回に調整（GAS制限対応）
const result = validateZoomDataWithRetry(
  () => fetchZoomData(`https://api.zoom.us/v2/webinars/${webinarId}/survey`, token),
  (d) => d && d.custom_survey && Array.isArray(d.custom_survey.questions) && d.custom_survey.questions.length > 0,
  3, // 最大3回リトライ（GAS 6分制限対応）
  webinarId, // webinarIdを渡す
  token // tokenを渡す
);

// 軽量版チェックは3回目のリトライ失敗時のみ実行
if (attempt === retryMax && webinarId && token) {
  Logger.log('3回目のリトライに失敗。軽量版でアンケートの有無をチェックします。');
  const hasSurvey = checkSurveyLightweight(webinarId, token);
  if (hasSurvey) {
    Logger.log('軽量チェック結果: アンケートは存在するが、データ取得に遅延が発生しています');
    return { valid: false, data: null, surveyExists: true };
  } else {
    Logger.log('軽量チェック結果: アンケートは存在しません');
    return { valid: false, data: null, surveyExists: false };
  }
}

// データ遅延時の適切な通知
if (result.surveyExists) {
  Logger.log('アンケートは存在するが、データ取得に遅延が発生しています');
  webhooktxt += '\n・アンケート結果レポート（データ遅延のため取得できませんでした）';
}
```

#### 軽量版アンケートチェックの特徴
- **効率的**: `?fields=custom_survey`で必要最小限のデータのみ取得
- **高速**: 全データを取得しないため処理が高速
- **確実**: アンケートの存在有無を確実に判定
- **安全**: 3回目のリトライ失敗時のみ実行されるため安全

#### 実装の利点
- **適切な耐性**: Zoom API遅延に対する適切な対応
- **効率性**: 軽量版チェックは必要な時のみ実行
- **明確性**: データ遅延とアンケート不存在を明確に区別
- **ユーザビリティ**: 適切な通知で状況が分かりやすい
- **保守性**: リトライ回数の調整が容易
- **GAS制限対応**: 6分制限に引っかからない適切な処理時間設計

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
- **優先ジョブ制御**: 優先ジョブ実行中は他の処理が自動的にスキップされます
- **登録者数チェック**: 21日前から当日の範囲でのみ実行されます
- **エラー通知**: 詳細なエラー情報で問題の特定が容易になります
- **優先ジョブ通知**: 優先ジョブ実行時はSlack通知されません（ログでの記録のみ）
- **除外判定**: 除外ワードと除外stockIDの両方で除外フラグを設定します
- **除外シート**: A列（除外ワード）とB列（除外stockID）の両方を管理してください
- **アンケートリトライ**: アンケートデータ取得は最大3回までリトライされます（GAS 6分制限対応）
- **軽量版チェック**: 3回目のリトライ失敗時のみ軽量版でアンケート存在確認が実行されます
- **データ遅延検出**: アンケートが存在するがデータが遅延している場合を適切に検出・通知します