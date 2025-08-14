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

## セットアップ手順
1. スクリプトプロパティに必要な値を設定
2. 各機能のテスト実行
3. 本格運用開始

## 注意事項
- 機密情報はすべてスクリプトプロパティで管理
- コードに直接機密情報を記載しない
- 環境別設定が可能