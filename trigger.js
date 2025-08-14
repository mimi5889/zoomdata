function checkTime() {// 8時台～20時台の場合は処理を実行しない
  // 現在の時刻を取得
  const now = new Date();

  // 日本時間に変換
  const JST_OFFSET = 9; // UTCからのオフセット（日本時間はUTC+9）
  const jstNow = new Date(now.getTime() + JST_OFFSET * 60 * 60 * 1000);

  // 時間を取得
  const hour = jstNow.getUTCHours();

  // 8時台～20時台の範囲を確認
  if (hour < 8 || hour > 20) {
    console.log("時間外:"+now);
    return; // 処理を終了
  }
  getMail();
}