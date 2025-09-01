/**
 * ユ別に登録されている名前が、BP情報一覧に登録されているかをチェックし、
 * 見つからない名前をログとGoogle Chatに出力するスクリプトです。
 *
 * 以下の条件を満たす行をチェック対象とします。
 * ・5列目（所属）に値が入力されている行
 * ・2列目（名前）が「作業者名」や「社員数」から始まらない行
 *
 * BP情報一覧に名前がなかった場合は、
 * ユ別の案件名、名前、顧客、所属をログとGoogle Chatに出力します。
 */

function checkNamesInSheets() {
  // =========================================================================
  // 設定項目：以下の情報を、ご自身のスプレッドシートに合わせて変更してください。
  // =========================================================================
  // ユ別の格納フォルダID
  const folderIdA = PropertiesService.getScriptProperties().getProperty('YUBETSU_FOLDER_ID');
  // ユ別のファイル名サフィックス
  const departmentSuffixes = ['ﾃﾞｼﾞﾀﾙ推進部', '業務推進部'];
  // ユ別のシート名
  const sheetNameA = '売上・支払情報'; 
  // ユ別の案件名が記載されている列番号
  const caseNameColumnA = 1; 
  // ユ別の個人名が記載されている列番号（A列は1, B列は2...）
  const nameColumnA = 2; 
  // ユ別の顧客が記載されている列番号
  const customerColumnA = 3; 
  // ユ別の所属が記載されている列番号
  const departmentColumnA = 5;

  // 実行日の前月を計算し、ファイル名を動的に生成
  const today = new Date();
  // 前月の1日を計算
  const prevMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const year = prevMonth.getFullYear();
  // 月を2桁表示にフォーマット
  const month = (prevMonth.getMonth() + 1).toString().padStart(2, '0');
  const fileNamePrefix = `${year}.${month}_`;

  // BP情報一覧のファイルID
  const spreadsheetIdB = PropertiesService.getScriptProperties().getProperty('BPICHIRAN_FILE_ID');
  // BP情報一覧のシート名
  const sheetNameB = 'フォームの回答 2';
  // BP情報一覧の個人名が記載されている列番号（A列は1, B列は2...）
  const nameColumnB = 6;

  // Google ChatのWebhook URL
  // Google Chatのスペースの「Webhookを作成」で取得したURLを設定
  // チャットスペース：協力会社委員会　（委員会メンバーの内部チャット）
  const chatWebhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_INNER_WEBHOOKURL');
  // チャットスペース：協力会社委員会　（FI本部幹部通知用チャット）
  //   const chatWebhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_OUTER_WEBHOOKURL');


  // =========================================================================
  // ここから下のコードは変更しないでください。
  // =========================================================================

  try {
    const folder = DriveApp.getFolderById(folderIdA);
    const filesToProcess = [];

    // 指定されたファイル名でフォルダ内を検索し、見つかったファイルをリストに追加
    for (const suffix of departmentSuffixes) {
      const fileName = `${fileNamePrefix}${suffix}`;
      const fileIterator = folder.getFilesByName(fileName);
      if (fileIterator.hasNext()) {
        filesToProcess.push(fileIterator.next());
        Logger.log(`ファイル「${fileName}」が見つかりました。`);
      } else {
        Logger.log(`警告：ファイル「${fileName}」が見つかりませんでした。`);
      }
    }

    if (filesToProcess.length === 0) {
      const errorMessage = `エラー：対象となるファイルが見つかりませんでした。スクリプトを終了します。`;
      Logger.log(errorMessage);
      sendToChat(chatWebhookUrl, errorMessage);
      return;
    }

    // BP情報一覧の名前をSetに格納して高速な検索を可能にする
    const ssB = SpreadsheetApp.openById(spreadsheetIdB);
    const sheetB = ssB.getSheetByName(sheetNameB);
    if (!sheetB) {
      const errorMessage = `エラー：BP情報一覧に「${sheetNameB}」というシートが見つかりません。`;
      Logger.log(errorMessage);
      sendToChat(chatWebhookUrl, errorMessage);
      return;
    }
    const valuesB = sheetB.getDataRange().getValues();

    const namesInB = new Set();
    // ヘッダー行をスキップするために i = 1 から開始
    for (let i = 1; i < valuesB.length; i++) {
      const name = valuesB[i][nameColumnB - 1];
      if (name) {
        // 名前からすべてのスペース（全角・半角）を削除してSetに追加
        namesInB.add(name.toString().trim().replace(/ |　/g, ''));
      }
    }
    
    let totalMissingCount = 0;
    const missingNames = [];
    
    // 見つかった各ファイルを処理
    for (const file of filesToProcess) {
      Logger.log(`\n--- ${file.getName()} の名前をチェック中 ---`);
      
      // DriveAppで取得したファイルオブジェクトからIDを取得し、スプレッドシートを開く
      const ssA = SpreadsheetApp.openById(file.getId());
      const sheetA = ssA.getSheetByName(sheetNameA);
      if (!sheetA) {
        const errorMessage = `エラー：ファイル「${file.getName()}」に「${sheetNameA}」というシートが見つかりません。`;
        Logger.log(errorMessage);
        sendToChat(chatWebhookUrl, errorMessage);
        continue; // 次のファイルへ
      }
      const valuesA = sheetA.getDataRange().getValues();

      let missingCount = 0;
      // ヘッダー行をスキップするために i = 1 から開始
      for (let i = 1; i < valuesA.length; i++) {
        const name = valuesA[i][nameColumnA - 1];
        const department = valuesA[i][departmentColumnA - 1];
        const caseName = valuesA[i][caseNameColumnA - 1];
        const customer = valuesA[i][customerColumnA - 1];

        // 5列目（所属）に値があり、名前も存在し、「作業者名」「社員数」から始まらない場合にチェック
        if (department && name) {
          const nameString = name.toString().trim();
          if (nameString.startsWith('作業者名') || nameString.startsWith('社員数')) {
            continue; // 対象外の行はスキップ
          }
          
          // 比較する名前からすべてのスペース（全角・半角）を削除
          const normalizedName = nameString.replace(/ |　/g, '');
          if (!namesInB.has(normalizedName)) {
            const logMessage = `  見つかりませんでした: 案件名「${caseName}」/ 名前「${name}」/ 顧客「${customer}」/ 所属「${department}」`;
            Logger.log(logMessage);
            missingNames.push({ caseName, name, customer, department });
            missingCount++;
          }
        }
      }
      totalMissingCount += missingCount;
      Logger.log(`完了：${file.getName()} から合計${missingCount}件の名前が、BP情報一覧に見つかりませんでした。`);
    }

    const chatMessageHeader = `${year}年${month}月のユ別に記載されている要員が、すべてBP情報一覧に登録済であることのチェックを行います。`;
    let chatMessageBody = '';

    if (totalMissingCount > 0) {
      chatMessageBody = `チェック結果、未登録は${totalMissingCount}名でした。\n\n`;
      missingNames.forEach(item => {
        chatMessageBody += `未登録: 案件名「${item.caseName}」/ 名前「${item.name}」/ 顧客「${item.customer}」/ 所属「${item.department}」\n`;
      });
    } else {
      chatMessageBody = 'チェック結果、すべての名前がBP情報一覧に登録されていました。';
    }

    Logger.log(`\n--- 全体の結果 ---`);
    Logger.log(chatMessageBody);
    sendToChat(chatWebhookUrl, `${chatMessageHeader}\n\n${chatMessageBody}`);

  } catch (e) {
    const errorMessage = `スクリプト実行中にエラーが発生しました：${e.toString()}`;
    Logger.log(errorMessage);
    sendToChat(chatWebhookUrl, errorMessage);
  }
}

/**
 * Google ChatのWebhook URLにメッセージを送信するヘルパー関数
 * @param {string} url - Google ChatのWebhook URL
 * @param {string} message - 送信するメッセージ
 */
function sendToChat(url, message) {
  try {
    const payload = JSON.stringify({ 'text': message });
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': payload,
    };
    UrlFetchApp.fetch(url, options);
    Logger.log('Google Chatにメッセージを送信しました。');
  } catch (e) {
    Logger.log(`Google Chatへのメッセージ送信中にエラーが発生しました：${e.toString()}`);
  }
}

/**
 * Google Chatに固定メッセージを投稿する関数。
 * Apps Scriptのトリガーで実行するように設定する。
 * @author T.Yoshida
 */
function postMessageToChat() {
  // チャットスペース：協力会社委員会　（委員会メンバーの内部チャット）
  const WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('CHAT_INNER_WEBHOOKURL');

  // 送信する固定メッセージ
  message = "ユ別をメールで受領していたら、10日までに、以下のフォルダにスプレッドシート形式で格納してください。 \n"
  message += "https://drive.google.com/drive/folders/" & PropertiesService.getScriptProperties().getProperty('YUBETSU_FOLDER_ID');

  const MESSAGES = {
    "text": message
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


