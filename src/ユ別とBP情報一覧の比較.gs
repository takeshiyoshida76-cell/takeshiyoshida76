/**
 * ユ別に登録されている名前が、BP情報一覧に登録されているかをチェックし、
 * 未登録の名前を管理者単位および顧客・案件単位でグループ化してログとGoogle Chatに出力する。
 * 管理者マスタで顧客名のみのレコードにも対応し、案件名が未指定の場合にフォールバックとして使用。
 * @author T.Yoshida
 * @throws {Error} スプレッドシートやフォルダが見つからない場合、またはGoogle Chat送信時にエラーが発生した場合
 */
function checkNamesInSheets() {
  // =========================================================================
  // 設定項目：スプレッドシートやファイルの情報を設定
  // =========================================================================
  // ユ別の格納フォルダID
  const folderIdA = PropertiesService.getScriptProperties().getProperty('YUBETSU_FOLDER_ID');
  // ユ別のファイル名サフィックス
  const departmentSuffixes = ['ﾃﾞｼﾞﾀﾙ推進部', '業務推進部'];
  // ユ別のシート名
  const sheetNameA = '売上・支払情報'; 
  // ユ別の案件名が記載されている列番号（A列=1）
  const caseNameColumnA = 1; 
  // ユ別の個人名が記載されている列番号
  const nameColumnA = 2; 
  // ユ別の顧客が記載されている列番号
  const customerColumnA = 3; 
  // ユ別の所属が記載されている列番号
  const departmentColumnA = 5;

  // 管理者マスタのスプレッドシートID
  const adminMasterId = PropertiesService.getScriptProperties().getProperty('ADMIN_MASTER_FILE_ID');
  // 管理者マスタのシート名
  const adminSheetName = 'シート1';
  // 管理者マスタの列番号（顧客名、案件名、管理者氏名）
  const adminCustomerColumn = 1;
  const adminCaseNameColumn = 2;
  const adminNameColumn = 3;

  // 実行日の前月を計算し、ファイル名を動的に生成
  const today = new Date();
  const prevMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const year = prevMonth.getFullYear();
  const month = (prevMonth.getMonth() + 1).toString().padStart(2, '0');
  const fileNamePrefix = `${year}.${month}_`;

  // BP情報一覧のファイルID
  const spreadsheetIdB = PropertiesService.getScriptProperties().getProperty('BPICHIRAN_FILE_ID');
  // BP情報一覧のシート名
  const sheetNameB = 'フォームの回答 1';
  // BP情報一覧の個人名が記載されている列番号
  const nameColumnB = 6;

  // Google ChatのWebhook URL（内部チャット用）
  const chatWebhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_INNER_WEBHOOKURL');

  // =========================================================================
  // メイン処理
  // =========================================================================
  try {
    // 管理者マスタを読み込み、顧客名＋案件名および顧客名のみをキーとする管理者マップを作成
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
    for (let i = 1; i < adminValues.length; i++) { // ヘッダー行をスキップ
      const customer = adminValues[i][adminCustomerColumn - 1]?.toString().trim();
      const caseName = adminValues[i][adminCaseNameColumn - 1]?.toString().trim() || '';
      const adminName = adminValues[i][adminNameColumn - 1]?.toString().trim();
      if (customer && adminName) {
        // 顧客名＋案件名（案件名が空の場合は顧客名のみ）のキーで管理者を登録
        const key = caseName ? `${customer}|${caseName}` : `${customer}|`;
        adminMap.set(key, adminName);
        Logger.log(`管理者マスタ登録: キー=${key}, 管理者=${adminName}`);
      } else {
        Logger.log(`警告：管理者マスタの行${i + 1}に不正なデータ（顧客名=${customer}, 管理者=${adminName}）をスキップ`);
      }
    }

    // ユ別ファイルを取得
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
      const errorMessage = `エラー：対象となるファイルが見つかりませんでした。スクリプトを終了します。`;
      Logger.log(errorMessage);
      sendToChat(chatWebhookUrl, errorMessage);
      return;
    }

    // BP情報一覧の名前をSetに格納（高速検索用）
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
        namesInB.add(name.toString().trim().replace(/ |　/g, '')); // スペースを除去して正規化
      }
    }
    
    let totalMissingCount = 0;
    const missingNamesByAdmin = new Map(); // 管理者ごとの未登録情報（顧客＋案件単位でサブグループ）
    const missingNamesNoAdmin = new Map(); // 管理者不明の未登録情報（顧客＋案件単位でサブグループ）

    // ユ別ファイルごとに名前をチェック
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

      let missingCount = 0;
      for (let i = 1; i < valuesA.length; i++) { // ヘッダー行をスキップ
        const name = valuesA[i][nameColumnA - 1];
        const department = valuesA[i][departmentColumnA - 1];
        const caseName = valuesA[i][caseNameColumnA - 1] || ''; // 空値を明示的に処理
        const customer = valuesA[i][customerColumnA - 1] || ''; // 空値を明示的に処理

        // 所属と名前が入力されており、対象外の名前でない場合にチェック
        if (department && name && customer) { // 顧客名が必須
          const nameString = name.toString().trim();
          if (nameString.startsWith('作業者名') || nameString.startsWith('社員数')) {
            continue;
          }
          
          const normalizedName = nameString.replace(/ |　/g, '');
          if (!namesInB.has(normalizedName)) {
            const logMessage = `  見つかりませんでした: 案件名「${caseName}」/ 名前「${name}」/ 顧客「${customer}」/ 所属「${department}」`;
            Logger.log(logMessage);
            
            // 管理者を特定（顧客名＋案件名を優先、なければ顧客名のみで検索）
            const adminKey = `${customer}|${caseName}`;
            let adminName = adminMap.get(adminKey);
            if (!adminName) {
              adminName = adminMap.get(`${customer}|`) || '管理者不明';
            }
            
            // 顧客＋案件をキーとするサブグループを作成
            const subKey = `${customer}|${caseName}`;
            const targetMap = adminName === '管理者不明' ? missingNamesNoAdmin : missingNamesByAdmin;
            
            if (!targetMap.has(adminName)) {
              targetMap.set(adminName, new Map());
            }
            const adminSubMap = targetMap.get(adminName);
            if (!adminSubMap.has(subKey)) {
              adminSubMap.set(subKey, { customer, caseName, items: [] });
            }
            adminSubMap.get(subKey).items.push({ name, department });
            missingCount++;
          }
        } else {
          Logger.log(`警告：ユ別ファイル「${file.getName()}」の行${i + 1}に不正なデータ（名前=${name}, 所属=${department}, 顧客=${customer}）をスキップ`);
        }
      }
      totalMissingCount += missingCount;
      Logger.log(`完了：${file.getName()} から合計${missingCount}件の名前が、BP情報一覧に見つかりませんでした。`);
    }

    // 通知メッセージを構築
    const chatMessageHeader = `@all ${year}年${month}月のユ別に記載されている要員が、すべてBP情報一覧に登録済であることのチェックを行いました。以下の未登録者が見つかりました。管理者の方は対応をお願いします。`;
    let chatMessageBody = '';

    if (totalMissingCount > 0) {
      chatMessageBody = `チェック結果、未登録は${totalMissingCount}名でした。\n\n`;

      // 管理者ごとの未登録情報を出力
      for (const [adminName, adminSubMap] of missingNamesByAdmin) {
        chatMessageBody += `【${adminName} 様 担当】\n`;
        Logger.log(`管理者=${adminName} のデータを処理中`);
        for (const [subKey, data] of adminSubMap) {
          const { customer, caseName, items } = data;
          if (!items || !Array.isArray(items)) {
            Logger.log(`エラー：キー=${subKey} のデータにitemsが不正（items=${items}）`);
            continue; // itemsが不正な場合はスキップ
          }
          chatMessageBody += `顧客: ${customer} / 案件名: ${caseName}\n`;
          items.forEach(item => {
            chatMessageBody += `  - 名前: ${item.name} / 所属: ${item.department}\n`;
          });
          chatMessageBody += `\n`; // 顧客＋案件ごとの空行
        }
        chatMessageBody += `\n`; // 管理者ごとの空行
      }

      // 管理者不明の未登録情報を出力
      if (missingNamesNoAdmin.size > 0) {
        chatMessageBody += `【管理者不明】\n`;
        Logger.log(`管理者不明のデータを処理中`);
        for (const [subKey, data] of missingNamesNoAdmin) {
          const { customer, caseName, items } = data;
          if (!items || !Array.isArray(items)) {
            Logger.log(`エラー：キー=${subKey} のデータにitemsが不正（items=${items}）`);
            continue; // itemsが不正な場合はスキップ
          }
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
