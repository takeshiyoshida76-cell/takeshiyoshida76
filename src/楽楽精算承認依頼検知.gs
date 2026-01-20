function notifyRakurakuToChat() {
  // 1. スクリプトプロパティからすべての設定を一括取得
  const props = PropertiesService.getScriptProperties().getProperties();
  const WEBHOOK_URL = props.WEBHOOK_URL;
  const RAKURAKU_FROM = props.RAKURAKU_FROM;
  const LABEL_NAME = props.LABEL_NAME;

  // 設定漏れチェック
  if (!WEBHOOK_URL || !RAKURAKU_FROM || !LABEL_NAME) {
    console.error("エラー: スクリプトプロパティ(WEBHOOK_URL, RAKURAKU_FROM, LABEL_NAME)をすべて設定してください。");
    return;
  }

  // ラベルの取得または作成
  const label = GmailApp.getUserLabelByName(LABEL_NAME) || GmailApp.createLabel(LABEL_NAME);
  
  // 検索クエリの組み立て
  const SEARCH_QUERY = `from:${RAKURAKU_FROM} "承認依頼" -label:${LABEL_NAME}`;
  const threads = GmailApp.search(SEARCH_QUERY);
  
  if (threads.length === 0) {
    console.log("新規の承認依頼メールはありません。");
    return;
  }

  threads.forEach(thread => {
    const messages = thread.getMessages();

    messages.forEach(message => {
      // 処理済みラベルが既にスレッドに付いているか再確認（二重送信防止）
      const currentLabels = thread.getLabels().map(l => l.getName());
      if (currentLabels.includes(LABEL_NAME)) return;

      const body = message.getPlainBody();
      let targetText = "";
      const startIndex = body.indexOf("伝票No.：");
      
      if (startIndex !== -1) {
        targetText = body.substring(startIndex);
        // 「経路：」「距離：」の行を削除
        targetText = targetText.replace(/^(経路|距離)[\s　]*[:：].*(\n|\r\n)?/gm, "");
      } else {
        targetText = "伝票内容の解析に失敗しました。";
      }

      // 送信内容の構築
      const payloadText = `📢 *楽楽精算 承認依頼*\n\n${targetText}`;
      
      // ログ出力
      console.log(`[送信対象データ]\n${payloadText}`);

      const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify({ "text": payloadText })
      };

      try {
        UrlFetchApp.fetch(WEBHOOK_URL, options);
        console.log("チャット送信成功");
        
        // Google Chat APIのレートリミット回避
        Utilities.sleep(2000); 
        
      } catch (e) {
        console.error(`送信失敗: ${e.toString()}`);
      }
    });

    // 全てのメッセージを処理後にラベルを付与
    thread.addLabel(label);
  });
}
