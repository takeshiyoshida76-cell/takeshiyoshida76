/**
 * ユ別に登録されている名前が、BP情報一覧に登録されているかをチェックし、
 * 未登録の名前を管理者単位および顧客・案件単位でグループ化してログとGoogle Chatに出力する。
 * 管理者マスタで顧客名のみのレコードにも対応し、案件名が未指定の場合にフォールバックとして使用。
 * @author T.Yoshida
 * @throws {Error} スプレッドシートやフォルダが見つからない場合、またはGoogle Chat送信時にエラーが発生した場合
 */
function checkNamesInSheets() {
  // =========================================================================
  // 設定項目
  // =========================================================================
  const folderIdA = PropertiesService.getScriptProperties().getProperty('YUBETSU_FOLDER_ID');
  const departmentSuffixes = ['ﾃﾞｼﾞﾀﾙ推進部', '業務推進部'];
  const sheetNameA = '売上・支払情報'; 
  const caseNameColumnA = 1; 
  const nameColumnA = 2; 
  const customerColumnA = 3; 
  const departmentColumnA = 5;

  const adminMasterId = PropertiesService.getScriptProperties().getProperty('ADMIN_MASTER_FILE_ID');
  const adminSheetName = 'シート1';
  const adminCustomerColumn = 1;
  const adminCaseNameColumn = 2;
  const adminNameColumn = 3;

  const today = new Date();
  const day = today.getDate();
  let targetMonth;
  if (day <= 9) {
    targetMonth = new Date(today.getFullYear(), today.getMonth() - 2, 1);
  } else {
    targetMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  }
  const year = targetMonth.getFullYear();
  const month = (targetMonth.getMonth() + 1).toString().padStart(2, '0');
  const fileNamePrefix = `${year}.${month}_`;

  const spreadsheetIdB = PropertiesService.getScriptProperties().getProperty('BPICHIRAN_FILE_ID');
  const sheetNameB = 'フォームの回答 1';
  const nameColumnB = 6;

  const chatWebhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_OUTER_WEBHOOKURL');

  // =========================================================================
  // メイン処理
  // =========================================================================
  try {
    // === 管理者マスタ読み込み ===
    const adminSs = SpreadsheetApp.openById(adminMasterId);
    const adminSheet = adminSs.getSheetByName(adminSheetName);
    if (!adminSheet) {
      const errorMessage = `エラー：管理者マスタに「${adminSheetName}」というシートが見つかりません。`;
      Logger.log(errorMessage);
      sendToChat(chatWebhookUrl, errorMessage);
      return;
    }
    const adminValues = adminSheet.getDataRange().getValues();
    const adminMap = new Map();
    for (let i = 1; i < adminValues.length; i++) {
      const customer = adminValues[i][adminCustomerColumn - 1]?.toString().trim();
      const caseName = adminValues[i][adminCaseNameColumn - 1]?.toString().trim() || '';
      const adminName = adminValues[i][adminNameColumn - 1]?.toString().trim();
      if (customer && adminName) {
        const key = caseName ? `${customer}|${caseName}` : `${customer}|`;
        adminMap.set(key, adminName);
      }
    }

    // === ユ別ファイル取得（変更なし）===
    const folder = DriveApp.getFolderById(folderIdA);
    const filesToProcess = [];
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
      const errorMessage = `エラー：対象となるファイルが見つかりませんでした。`;
      Logger.log(errorMessage);
      sendToChat(chatWebhookUrl, errorMessage);
      return;
    }

    // === BP情報一覧の名前セット ===
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
    for (let i = 1; i < valuesB.length; i++) {
      const name = valuesB[i][nameColumnB - 1];
      if (name) {
        namesInB.add(name.toString().trim().replace(/ |　/g, ''));
      }
    }

    let totalMissingCount = 0;
    const missingNamesByAdmin = new Map();
    const missingNamesNoAdmin = new Map();

    // === ユ別ファイルごとの処理 ===
    for (const file of filesToProcess) {
      Logger.log(`\n--- ${file.getName()} の名前をチェック中 ---`);
      const ssA = SpreadsheetApp.openById(file.getId());
      const sheetA = ssA.getSheetByName(sheetNameA);
      if (!sheetA) {
        const errorMessage = `エラー：ファイル「${file.getName()}」に「${sheetNameA}」というシートが見つかりません。`;
        Logger.log(errorMessage);
        sendToChat(chatWebhookUrl, errorMessage);
        continue;
      }
      const valuesA = sheetA.getDataRange().getValues();

      let lastCustomer = ''; // 顧客名を前方参照で保持
      let missingCount = 0;

      for (let i = 1; i < valuesA.length; i++) {
        let caseName = valuesA[i][caseNameColumnA - 1] || '';
        let rawName = valuesA[i][nameColumnA - 1];
        let rawCustomer = valuesA[i][customerColumnA - 1];
        let department = valuesA[i][departmentColumnA - 1];

        // 文字列化＆トリム
        caseName = caseName?.toString().trim() || '';
        const nameString = rawName?.toString().trim() || '';
        const customerString = rawCustomer?.toString().trim();

        // === 顧客名の前方参照処理 ===
        if (customerString && !/^\d+$/.test(customerString.replace(/,/g, ''))) {
          // 数字（金額）や空白でなければ、顧客名として更新
          lastCustomer = customerString;
        }
        // 最終的な顧客名（前方参照済み）
        const customer = lastCustomer;

        // チェック対象外の行はスキップ
        if (nameString.startsWith('作業者名') || nameString.startsWith('社員数') || !nameString || !department || !customer) {
          continue;
        }

        const normalizedName = nameString.replace(/ |　/g, '');
        if (!namesInB.has(normalizedName)) {
          Logger.log(`  未登録: 案件「${caseName}」/ 名前「${nameString}」/ 顧客「${customer}」`);

          // 管理者特定
          const adminKey = `${customer}|${caseName}`;
          let adminName = adminMap.get(adminKey);
          if (!adminName) {
            adminName = adminMap.get(`${customer}|`) || '管理者不明';
          }

          const subKey = `${customer}|${caseName}`;
          const targetMap = adminName === '管理者不明' ? missingNamesNoAdmin : missingNamesByAdmin;

          if (!targetMap.has(adminName)) {
            targetMap.set(adminName, new Map());
          }
          const adminSubMap = targetMap.get(adminName);
          if (!adminSubMap.has(subKey)) {
            adminSubMap.set(subKey, { customer, caseName, items: [] });
          }
          adminSubMap.get(subKey).items.push({ name: nameString, department });

          missingCount++;
        }
      }
      totalMissingCount += missingCount;
      Logger.log(`完了：${file.getName()} → ${missingCount}件未登録`);
    }

    // === 通知メッセージ構築 ===
    const chatMessageHeader = `@all ${year}年${month}月のユ別に記載されている要員が、すべてBP情報一覧に登録済であることのチェックを行いました。以下の未登録者が見つかりました。管理者の方は対応をお願いします。`;
    let chatMessageBody = '';

    if (totalMissingCount > 0) {
      chatMessageBody = `チェック結果、未登録は${totalMissingCount}名でした。\n\n担当が間違っている可能性がございます。間違っている場合は、委員会へ連絡してください。\n\n`;

      // 管理者ごと
      for (const [adminName, adminSubMap] of missingNamesByAdmin) {
        chatMessageBody += `【${adminName} 様 担当】\n`;
        for (const [subKey, data] of adminSubMap) {
          const { customer, caseName, items } = data;
          chatMessageBody += `顧客: ${customer} / 案件名: ${caseName}\n`;
          items.forEach(item => {
            chatMessageBody += `  - 名前: ${item.name} / 所属: ${item.department}\n`;
          });
          chatMessageBody += `\n`;
        }
        chatMessageBody += `\n`;
      }

      // 管理者不明
      if (missingNamesNoAdmin.size > 0) {
        chatMessageBody += `【管理者不明】\n`;
        for (const [subKey, data] of missingNamesNoAdmin) {
          const { customer, caseName, items } = data;
          chatMessageBody += `顧客: ${customer} / 案件名: ${caseName}\n`;
          items.forEach(item => {
            chatMessageBody += `  - 名前: ${item.name} / 所属: ${item.department}\n`;
          });
          chatMessageBody += `\n`;
        }
      }
    } else {
      chatMessageBody = 'チェック結果、すべての名前がBP情報一覧に登録されていました。';
    }

    Logger.log(`\n--- 全体の結果 ---\n${chatMessageBody}`);
    sendToChat(chatWebhookUrl, `${chatMessageHeader}\n\n${chatMessageBody}`);

  } catch (e) {
    const errorMessage = `スクリプト実行中にエラーが発生しました：${e.toString()}`;
    Logger.log(errorMessage);
    sendToChat(chatWebhookUrl, errorMessage);
  }
}

/**
 * Google ChatのWebhook URLにメッセージを送信するヘルパー関数。
 * @param {string} url - Google ChatのWebhook URL
 * @param {string} message - 送信するメッセージ
 * @author T.Yoshida
 * @throws {Error} Webhook送信時にエラーが発生した場合
 */
function sendToChat(url, message) {
  // JSONペイロードを作成
  const payload = JSON.stringify({ 'text': message });
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': payload,
  };

  try {
    UrlFetchApp.fetch(url, options);
    Logger.log('Google Chatにメッセージを送信しました。');
  } catch (e) {
    Logger.log(`Google Chatへのメッセージ送信中にエラーが発生しました：${e.toString()}`);
  }
}

/**
 * Google Chatに固定メッセージを投稿する関数。トリガーで定期実行される。
 * ユ別スプレッドシートの提出を促すメッセージを送信。
 * @author T.Yoshida
 * @throws {Error} Webhook送信時にエラーが発生した場合
 */
function postMessageToChat() {
  // Google ChatのWebhook URL（内部チャット用）
  const WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('CHAT_INNER_WEBHOOKURL');

  // 送信する固定メッセージ
  let message = "ユ別をメールで受領していたら、10日までに、以下のフォルダにスプレッドシート形式で格納してください。\n";
  message += `https://drive.google.com/drive/folders/${PropertiesService.getScriptProperties().getProperty('YUBETSU_FOLDER_ID')}`;

  const MESSAGES = {
    "text": message
  };

  // メッセージをJSON文字列に変換
  const payload = JSON.stringify(MESSAGES);

  // HTTPリクエストのオプションを設定
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": payload,
    "muteHttpExceptions": true // エラー時に例外を抑制
  };

  try {
    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    Logger.log("メッセージが正常に投稿されました。レスポンスコード：" + response.getResponseCode());
    Logger.log("レスポンス本文：" + response.getContentText());
  } catch (e) {
    Logger.log("メッセージの投稿中にエラーが発生しました：" + e.message);
  }
}
