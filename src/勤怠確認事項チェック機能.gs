/**
 * 勤怠確認事項チェック機能
 * 指定されたExcelマスターから勤怠確認事項を抽出し、対応が必要な項目をWebhookへ通知する。
 * 
 * 運用フロー:
 * 1. Excelマスターをスプレッドシート形式に一時変換
 * 2. 全シートを走査し、ステータスが「未対応」または「再確認」の行を抽出
 * 3. 一時ファイルを削除し、Webhookへ通知
 */
function runTaskChecker() {
  const props = PropertiesService.getScriptProperties();

  // --- 曜日・祝日判定の制御 ---
  const skipHolidays = props.getProperty('SKIP_HOLIDAYS') === 'true'; // プロパティを取得
  
  if (skipHolidays) {
    const today = new Date();
    const dayOfWeek = today.getDay();
    
    // 土日チェック
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      console.log("土日のためスキップします。");
      return;
    }
    
    // 祝日チェック
    const calendarId = "ja.japanese#holiday@group.v.calendar.google.com";
    const cal = CalendarApp.getCalendarById(calendarId);
    const events = cal.getEventsForDay(today);

    // イベントの説明文（description）に「祝日」という文字が含まれているかを確認
    const isHoliday = events.some(event => {
      const desc = event.getDescription(); // イベントの詳細説明文を取得
      return desc.includes("祝日");
    });

    if (isHoliday) {
      console.log("祝日のためスキップします: " + events[0].getTitle());
      return;
    }
  }

  const WEBHOOK_URL = props.getProperty('WEBHOOK_URL');
  const EXCEL_FILE_ID = props.getProperty('SPREADSHEET_ID');
  const TEMP_FOLDER_ID = props.getProperty('TEMP_FOLDER_ID');

  if (!WEBHOOK_URL || !EXCEL_FILE_ID || !TEMP_FOLDER_ID) {
    throw new Error('必要なスクリプトプロパティが設定されていません。');
  }

  // 1. マスターファイルを一時変換して読み込み
  const excelFile = DriveApp.getFileById(EXCEL_FILE_ID);
  const resource = {
    title: `AUTO_TEMP_${new Date().getTime()}`,
    parents: [{ id: TEMP_FOLDER_ID }],
    mimeType: MimeType.GOOGLE_SHEETS
  };

  // Drive APIを使用してExcelをスプレッドシートとして作成
  const tempSsFile = Drive.Files.create(resource, excelFile.getBlob());
  const ss = SpreadsheetApp.openById(tempSsFile.id);
  const sheets = ss.getSheets();

  // 2. 設定値定義
  const TARGET_COL_INDEX = 4; // E列：ステータス
  const TARGET_STATUSES = ['未対応', '再確認'];
  let message = "【勤怠確認事項：対応が必要です】\n";
  message += 'https://docs.google.com/spreadsheets/d/' + EXCEL_FILE_ID + '\n';
  let hasPending = false;

  // 3. 各シートのデータ走査
  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return; // ヘッダーのみの場合はスキップ

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[TARGET_COL_INDEX];

      if (status && TARGET_STATUSES.includes(status)) {
        hasPending = true;

        // 日付の整形（Date型であればM/d形式に変換）
        const dateValue = row[2];
        const formattedDate = (dateValue instanceof Date) 
          ? Utilities.formatDate(dateValue, "JST", "M/d") 
          : dateValue;

        message += `\n[${sheet.getName()}] ${row[1]} : ${formattedDate} : ${row[3]} : ${status}`;
      }
    }
  });

  // 4. クリーンアップ
  Drive.Files.remove(tempSsFile.id);

  // 5. 通知送信
  if (hasPending) {
    console.log("通知対象あり");
    postToWebhook(WEBHOOK_URL, message);
  } else {
    console.log("通知対象なし");
  }
}

/**
 * 指定されたURLへWebhook通知を行う
 * @param {string} url - 通知先のWebhook URL
 * @param {string} message - 送信するテキストメッセージ
 */
function postToWebhook(url, message) {
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({ "text": message })
  };
  UrlFetchApp.fetch(url, options);
}
