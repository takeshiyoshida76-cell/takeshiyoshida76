function onFormSubmit(e) {
  // フォームの全項目を取得
  const formResponses = e.response.getItemResponses();

  // 質問と回答をループして処理
  for (const response of formResponses) {
    // 質問のタイトルが「横展開可否」で始まるかを確認
    if (response.getItem().getTitle().startsWith('横展開可否')) {
      // 回答が「可」であるかを確認
      if (response.getResponse() === '可') {
        // 条件が満たされたら、メール送信などの処理を実行
        // ここでは例として onSendChat 関数を呼び出します
        onSendChat(e);
        
        // 目的の処理が完了したのでループを抜ける
        break; 
      }
    }
  }
  // 回答編集URLを編集
  addEditUrlsToExistingResponses();  
}

function sendEmail(formResponses) {
  var emailAddress = 'メールアドレス';
  var subject = '横展開「可」が入力されました';
  var message = '回答内容\n' + formResponses.join('\n');
  MailApp.sendEmail(emailAddress, subject, message);
}

/**
 * Googleフォームの内容をGoogle Chatに投稿する。
 * @author T.Yoshida
 */
function onSendChat(e) {

  // チャットスペース：協力会社委員会　（委員会メンバーの内部チャット）
  // const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAQAfPic_ks/messages?☆☆☆
  // チャットスペース：協力会社委員会　（FI本部幹部通知用チャット）
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAQAnWJ7tYU/messages?☆☆☆";
  
  // フォームの回答内容を格納する変数
  let message;

  // メッセージに含める質問のタイトルを配列で定義
  // ここに定義された文字列で始まる質問タイトルが対象となり、
  // この文字列がそのままメッセージのラベルとして使用されます。
  const targetQuestions = [
    "記入者氏名",
    "対象BPの氏名",
    "対象BPの自宅最寄り駅",
    "保有スキル",
    "終了予定時期"
  ];

  // eが定義されている（トリガーから実行された）か確認
  if (e && e.response) {
    // フォームイベントからメッセージを生成する場合
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();

    // メッセージの本文を生成
    message = "横展開「可」が入力されました！ 詳細はBP情報一覧をご確認お願いします。\n\n";

    // 各質問と回答をループして処理
    itemResponses.forEach(itemResponse => {
      const questionTitle = itemResponse.getItem().getTitle();

      // questionTitleがtargetQuestions配列のいずれかの項目で始まるかを確認
      const foundLabel = targetQuestions.find(label => questionTitle.startsWith(label));

      if (foundLabel) {
        const answer = itemResponse.getResponse();
        // 見つかったラベルをそのまま使用してメッセージに追加
        message += `${foundLabel}： ${answer}\n`;
      }
    });
  } else {
    // 手動実行の場合、またはeがない場合の代替メッセージ
    message = "【テスト】手動実行されました！\nこれはフォームの回答がない場合のテストメッセージです。\n";
  }

  message += "BP情報一覧のURL： https://docs.google.com/spreadsheets/d/☆☆☆";

  // メッセージをJSON形式に整形
  const payload = {
    "text": message
  };

  // HTTPリクエストのオプションを設定
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  try {
    // Webhook URLにHTTPリクエストを送信
    UrlFetchApp.fetch(WEBHOOK_URL, options);
    Logger.log("メッセージをGoogle Chatに送信しました。");
  } catch (err) {
    Logger.log("メッセージの送信に失敗しました: " + err.message);
  }
}

/**
 * Googleフォームの個別の回答編集URLをスプレッドシートに取得するスクリプト
 * @author T.Yoshida
 *
 * 既存の回答に対して、後から編集URLを一括で取得してスプレッドシートに追加する。
 *
 */
function addEditUrlsToExistingResponses() {
  const form = FormApp.getActiveForm();
  const responses = form.getResponses(); // すべての回答を取得
  
  const spreadsheet = SpreadsheetApp.openById("☆☆☆");
  const sheet = spreadsheet.getSheetByName("フォームの回答 2"); // 回答が記録されるシート名

  if (!sheet) {
    Logger.log("エラー: 'フォームの回答 2' という名前のシートが見つかりません。シート名を確認してください。");
    return;
  }

  // 編集URLを記録する列を特定
  // 例として、既存の最終列の隣に「編集URL」列を追加
  let editUrlColumn = sheet.getLastColumn() + 1;
  // もし既に「編集URL」というヘッダーがあれば、その列を使用
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existingEditUrlColumnIndex = headers.indexOf("編集URL");
  if (existingEditUrlColumnIndex !== -1) {
    editUrlColumn = existingEditUrlColumnIndex + 1; // 0-indexed to 1-indexed
  } else {
    // ヘッダーがない場合は追加
    sheet.getRange(1, editUrlColumn).setValue("編集URL");
  }

  // 各回答に対して編集URLを取得し、スプレッドシートに書き込む
  responses.forEach((response, index) => {
    // 回答が記録されている行は、スプレッドシートの2行目から始まるため、+2 する
    // (ヘッダー行が1行目、最初の回答が2行目)
    const row = index + 2; 
    const currentEditUrl = sheet.getRange(row, editUrlColumn).getValue();

    // 既にURLが書き込まれていない場合のみ処理
    if (!currentEditUrl) {
      const editUrl = response.getEditResponseUrl();
      sheet.getRange(row, editUrlColumn).setValue(editUrl);
      Logger.log(`既存の回答 (行 ${row}) に編集URLを追加しました: ${editUrl}`);
    }
  });

  Logger.log("既存の回答への編集URLの追加処理が完了しました。");
}

/**
 * Google Chatに固定メッセージを投稿する関数。
 * Apps Scriptのトリガーで実行するように設定する。
 * @author T.Yoshida
 */
function postMessageToChat() {
  // チャットスペース：協力会社委員会　（委員会メンバーの内部チャット）
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAQAfPic_ks/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=nvLYJglx6es6vDcTru30XroCT9gftUpq1rNwziKsLJ4"
  // チャットスペース：協力会社委員会　（FI本部幹部通知用チャット）
  // const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAQAnWJ7tYU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=cjOzt67ye8ZfC065Wf9iAiyONqVcM0kI_xdIqHcGt_8";

  // 送信する固定メッセージ
  message = "<users/all> \n"
  message += "協力会社管理委員会よりお知らせです。 \n"
  message += "BP情報一覧の更新にご協力ください。 \n"
  message += "　・登録済みの内容更新 \n"
  message += "　・新規参画者の追加 \n"
  message += "ユーザー別一覧が発信され次第、委員会で突合せ作業を実施いたします。 \n"
  message += "不明点や要望などあれば委員会メンバーもしくは本スペースにてご連絡ください。 \n\n"
  message += "●登録済み内容の更新（一覧に更新用URLがあります） \n"
  message += "BP情報一覧：https://docs.google.com/spreadsheets/d/1-aPpSVwhBAF4DoD1xeydhYDalPS6h6S_WQEtn83ofAw/edit?gid=1528605409#gid=1528605409 \n"
  message += "●新規登録URL \n"
  message += "新規入力フォーム：https://docs.google.com/forms/d/e/1FAIpQLSdat0RMcXEfde_nc-6hkD3C9vgG9P3zkywPu3hty48B4uyNDg/viewform \n"

  const MESSAGES = {
    "text": message,
    "annotations": [
      {
        "type": "USER_MENTION",
        "start_index": 0,
        "length": 11, // "<users/all>" の文字数
        "user_mention": {
          "user": {
            "name": "users/all",
            "type": "ALL"
          },
          "type": "MENTION"
        }
      }
    ]
  };
  // ===============================

  // メッセージをJSON文字列に変換
  const payload = JSON.stringify(MESSAGES);

  // ウェブフックにPOSTリクエストを送信
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": payload,
    "muteHttpExceptions": true // エラー時に例外を発生させず、レスポンスを返却
  };

  try {
    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    Logger.log("メッセージが正常に投稿されました。レスポンスコード：" + response.getResponseCode());
    Logger.log("レスポンス本文：" + response.getContentText());
  } catch (e) {
    Logger.log("メッセージの投稿中にエラーが発生しました：" + e.message);
  }
}

