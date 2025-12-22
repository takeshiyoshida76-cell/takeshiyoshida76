/**
 * Akashi勤怠チェック自動化スクリプト (Google Apps Script)
 * - 目的: Akashiの勤務表ページにアクセスし、出退勤の打刻漏れがないかチェックし、
 * Google Chat Webhook経由で通知を行います。
 * * * * 【仕様】:
 * 1. 勤怠サマリURLは日付なし (当日チェック) に固定。
 * 2. 出勤通知: 予定開始時刻の5分前以降、未打刻なら通知。
 * 3. 退勤通知: 予定終了時刻の15分後以降、未打刻なら通知。
 * 4. 理由欄 に '残業終了=HH:MM' の形式で残業終了時刻が記載されていた場合、
 * その HH:MM を新しい退勤予定時刻として採用し、チェックに使用する。
 * 5. 特定の社員を特定の日付でアラート対象外にする設定をスプレッドシートで管理。
 *    - 除外設定された社員は、出勤アラートと退勤アラートの両方が抑制される。
 */

// ==============================================================================
// 1. 定数設定
// ==============================================================================

// 通知対象の氏名を設定 (空の場合は全員が対象)。名字のみでOKです。（例: ['葭田', '南']）
const TARGET_NAMES = []; 
const LOGIN_URL = 'https://atnd.ak4.jp/ja/login';
const MANAGER_URL = 'https://atnd.ak4.jp/ja/manager';
const ROOT_JA_URL = 'https://atnd.ak4.jp/ja'; 
const CURRENT_ATTENDANCE_URL = 'https://atnd.ak4.jp/ja/manager/current_attendance_status';
const ATTENDANCE_URL = 'https://atnd.ak4.jp/ja/manager/daily_summary'; 
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36';
const EXCLUSION_SHEET_ID = PropertiesService.getScriptProperties().getProperty('EXCLUSION_SHEET_ID'); // 除外設定スプレッドシートのID
const EXCLUSION_SHEET_NAME = '除外設定'; // 除外設定シートの名前

// ==============================================================================
// 2. ユーティリティ関数
// ==============================================================================

/**
 * スクリプトプロパティから認証情報を取得
 * @description スクリプトプロパティから会社ID、ログインID、パスワード、Webhook URL、除外設定シートIDを取得します。
 * @returns {Object} 認証情報オブジェクト（companyId, loginId, password, webhookUrl, exclusionSheetId）
 * @throws {Error} スクリプトプロパティが未設定の場合、エラーがログに記録される
 */
function getCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    companyId: scriptProperties.getProperty('COMPANY_ID'),
    loginId: scriptProperties.getProperty('LOGIN_ID'),
    password: scriptProperties.getProperty('PASSWORD'),
    webhookUrl: scriptProperties.getProperty('WEBHOOK_URL'),
    exclusionSheetId: scriptProperties.getProperty('EXCLUSION_SHEET_ID')
  };
}

/**
 * 現在のJST日時を取得
 * @returns {{date: Date, dateString: string}} JST DateオブジェクトとYYYYMMDD形式の文字列
 */
function getCurrentJSTAndDateString() {
  const now = new Date();
  const jstTimeZone = 'Asia/Tokyo';
  const nowJST = new Date(Utilities.formatDate(now, jstTimeZone, 'yyyy-MM-dd HH:mm:ss') + '+09:00');
  const dateString = Utilities.formatDate(nowJST, jstTimeZone, 'yyyyMMdd');
  return { date: nowJST, dateString: dateString };
}

/**
 * HH:MM形式の時刻文字列を、指定した基準日に基づくDateオブジェクトに変換する
 */
function parseTime(timeStr, baseDate) {
  const [hours, minutes] = timeStr.split(':').map(Number);
  const date = new Date(baseDate);
  date.setHours(hours, minutes, 0, 0);
  return date;
}

/**
 * ランダムな遅延を追加
 * @description 指定した範囲（ミリ秒）でランダムな遅延を挿入し、サーバー負荷を軽減します。
 * @param {number} min 最小遅延時間（ミリ秒）
 * @param {number} max 最大遅延時間（ミリ秒）
 */
function addRandomDelay(min, max) {
  const delay = Math.floor(Math.random() * (max - min + 1)) + min;
  Utilities.sleep(delay);
}

/**
 * Set-CookieヘッダーからCookieマップを更新
 * @description HTTPレスポンスのSet-CookieヘッダーからCookieマップを更新します。
 * @param {Object} cookieMap 現在のCookieマップ（キー:値ペア）
 * @param {string|string[]} setCookieHeaders Set-Cookieヘッダー（単一または配列）
 * @returns {Object} 更新されたCookieマップ
 */
function updateCookieMap(cookieMap, setCookieHeaders) {
  if (!setCookieHeaders) return cookieMap;
  const cookieArray = Array.isArray(setCookieHeaders) ? setCookieHeaders : [setCookieHeaders];
  cookieArray.forEach(cookieStr => {
    const pair = cookieStr.split(';')[0].trim();
    if (pair.includes('=')) {
      const [key, value] = pair.split('=').map(s => s.trim());
      cookieMap[key] = value;
    }
  });
  return cookieMap;
}

/**
 * CookieマップをHTTPリクエストヘッダー用の文字列に変換
 * @description Cookieマップを「key=value; key=value」形式の文字列に変換します。
 * @param {Object} cookieMap Cookieのキー:値ペア
 * @returns {string} HTTPリクエスト用のCookieヘッダー文字列
 */
function formatCookieHeader(cookieMap) {
  return Object.entries(cookieMap)
    .map(([key, value]) => `${key}=${value}`)
    .join('; ');
}

/**
 * 勤務ステータス文字列から休暇の種類を判定
 * @param {string} statusText Akashiから取得したステータス文字列
 * @returns {string|null} "休暇", "午前休", "午後休" または null
 */
function getLeaveType(statusText) {
  if (!statusText) return null;
  if (statusText.match(/午前半年休/)) return "午前休";
  if (statusText.match(/午後半年休/)) return "午後休";
  if (statusText.match(/年休|振替休日|代休|記念日休暇|忌引|産休/)) return "休暇";
  return null;
}

/**
 * 除外設定スプレッドシートから対象者を取得
 * @description 指定された日付の除外対象者リストをスプレッドシートから取得します。
 * @param {string} dateString YYYYMMDD形式の日付
 * @returns {string[]} 除外対象者の名字リスト
 */
function getExcludedNames(dateString) {
  const credentials = getCredentials();
  const sheetId = credentials.exclusionSheetId;
  if (!sheetId) {
    Logger.log('エラー: EXCLUSION_SHEET_IDが設定されていません。除外設定をスキップします。');
    return [];
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(EXCLUSION_SHEET_NAME);
    if (!sheet) {
      Logger.log(`エラー: シート「${EXCLUSION_SHEET_NAME}」が見つかりません。除外設定をスキップします。`);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const excludedNames = [];

    // ヘッダー行をスキップし、指定された日付の対象者を抽出
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0]?.toString().trim();
      const name = data[i][1]?.toString().trim();
      if (rowDate === dateString && name) {
        excludedNames.push(name);
      }
    }

    Logger.log(`情報: ${dateString} の除外対象者: ${excludedNames.length > 0 ? excludedNames.join(', ') : 'なし'}`);
    return excludedNames;
  } catch (e) {
    Logger.log(`エラー: 除外設定シートの読み込みに失敗: ${e.message}`);
    return [];
  }
}

// ==============================================================================
// 3. メイン処理 (エントリーポイント)
// ==============================================================================

/**
 * 勤怠チェックのメイン処理
 * @description Akashiにログインし、勤怠データをチェックしてGoogle Chatに通知します。ログイン失敗時はリトライを行い、失敗した場合にエラー通知を送信します。
 */
function main() {
  Logger.log('お知らせ: 実行開始');
  const sessionCookie = loginToAkashiWithRetry();

  if (sessionCookie) {
    checkAttendance(sessionCookie);
  } else {
    const jstInfo = getCurrentJSTAndDateString();
    const dateFormatted = Utilities.formatDate(jstInfo.date, 'Asia/Tokyo', 'yyyy年MM月dd日');
    const errorMessage = `【勤怠チェック失敗】${dateFormatted}: ログインに失敗したため、勤怠チェックができませんでした。`;
    sendToGoogleChat(errorMessage, false);
    Logger.log('エラー: ログインに失敗したため、処理を中断します。');
  }
  Logger.log('お知らせ: 実行完了');
}

/**
 * 勤怠データのチェックと通知
 * @description Akashiの勤怠ページからデータを取得し、打刻漏れをチェックしてGoogle Chatに通知します。
 * @param {string} sessionCookie AkashiのセッションCookie
 */
function checkAttendance(sessionCookie) {
  const jstInfo = getCurrentJSTAndDateString();
  const now = jstInfo.date;
  const todayString = jstInfo.dateString; // YYYYMMDD形式の今日の日付
  
  Logger.log(`情報: 対象とする勤怠サマリURL (当日): ${ATTENDANCE_URL}`);
  
  const cookieMap = sessionCookie.split('; ').reduce((acc, pair) => {
    const [key, value] = pair.split('=').map(s => s.trim());
    if (key) acc[key] = value;
    return acc;
  }, {});
  
  const urls = [MANAGER_URL, CURRENT_ATTENDANCE_URL, ATTENDANCE_URL];
  let currentCookieMap = cookieMap;
  let htmlContent = '';

  for (const url of urls) {
    const fetchResult = fetchPageHTML(currentCookieMap, url);
    if (!fetchResult.html) {
      Logger.log(`エラー: ${url} の取得に失敗したため、処理を中断します。`);
      return;
    }
    htmlContent = fetchResult.html;
    currentCookieMap = fetchResult.cookieMap;
    addRandomDelay(2000, 5000);
  }

  if (!htmlContent) {
    Logger.log('エラー: 勤怠HTMLの取得に失敗したため、処理を中断します。');
    return;
  }
  
  const employeeData = parseAttendanceHTML(htmlContent);
  const messages = [];
  const excludedNames = getExcludedNames(todayString); // 除外対象者リストを取得

  Logger.log('情報: 抽出された全社員数: ' + employeeData.length);

  employeeData.forEach(employee => {
    const lastName = employee.name.trim().split(/\s+/)[0];
    if (TARGET_NAMES.length === 0 || TARGET_NAMES.includes(lastName)) {
      const nowTime = now.getTime();
      Logger.log('チェック情報: 氏名:' + employee.name + ' 予定開始:' + employee.scheduledStart + ' 予定終了:' + employee.scheduledEnd + ' 打刻開始:' + employee.stampStart + ' 打刻終了:' + employee.stampEnd);

      // スプレッドシートで指定された日付と名字が一致しない場合にアラートを送信
      if (!excludedNames.includes(lastName)) {
        if (employee.scheduledStart !== '--:--') {
          const scheduledStartTime = parseTime(employee.scheduledStart, now);
          const fiveMinutesBefore = scheduledStartTime.getTime() - 5 * 60 * 1000;
          if (employee.stampStart === '--:--' && nowTime >= fiveMinutesBefore) {
            messages.push(`${employee.name} さん、予定出勤時刻 ${employee.scheduledStart} の出勤打刻が行われていません。至急、打刻してください。`);
          }
        }

        if (employee.scheduledEnd !== '--:--') {
          const scheduledEndTime = parseTime(employee.scheduledEnd, now);
          const fifteenMinutesAfter = scheduledEndTime.getTime() + 15 * 60 * 1000;
          if (employee.stampEnd === '--:--' && nowTime >= fifteenMinutesAfter) {
            messages.push(`${employee.name} さん、予定退勤時刻 ${employee.scheduledEnd} の退勤打刻が行われていません。至急、打刻してください。`);
          }
        }
      }
    }
  });

  if (messages.length > 0) {
    const dateFormatted = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日');
    const header = `【打刻アラート】${dateFormatted} の打刻漏れが${messages.length}件あります！\n`;
    const messageText = header + messages.join('\n');
    sendToGoogleChat(messageText, true);
    Logger.log('情報: 送信したメッセージ:\n' + messageText);
  } else {
    Logger.log('情報: 送信するメッセージはありません。');
  }
}

/**
 * ログイン失敗時にリトライを行うラッパー関数
 * @returns {string|null} セッションCookie（成功時）またはnull（失敗時）
 */
function loginToAkashiWithRetry() {
  const maxRetries = 3;
  let retryCount = 0;

  while (retryCount < maxRetries) {
    Logger.log(`情報: ログイン試行 (${retryCount + 1}/${maxRetries})`);
    const sessionCookie = loginToAkashi();
    if (sessionCookie) {
      return sessionCookie;
    }
    retryCount++;
    Logger.log(`警告: ログイン失敗、リトライします（試行 ${retryCount}/${maxRetries}）`);
    addRandomDelay(1000, 2000);
  }

  Logger.log(`エラー: 最大リトライ回数（${maxRetries}）を超過、ログイン失敗`);
  return null;
}

/**
 * Akashiにログインし、成功した場合にセッションCookieを返す
 */
function loginToAkashi() {
  Logger.log('お知らせ: ログイン処理開始');
  const credentials = getCredentials();
  if (!credentials.companyId || !credentials.loginId || !credentials.password) {
    Logger.log('エラー: 認証情報が不足しています。スクリプトプロパティを確認してください。');
    return null;
  }

  let cookieMap = {};
  const getOptions = {
    method: 'GET',
    muteHttpExceptions: true,
    headers: {
      'User-Agent': USER_AGENT,
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
      'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
      'Upgrade-Insecure-Requests': '1',
      'DNT': '1'
    }
  };
  
  try {
    const getResponse = UrlFetchApp.fetch(LOGIN_URL, getOptions);
    const html = getResponse.getContentText('UTF-8');
    const headers = getResponse.getAllHeaders();
    const match = html.match(/<input[^>]*name="authenticity_token"[^>]*value="([^"]*)"/);
    if (!match || !match[1]) {
      Logger.log('エラー: CSRFトークンが見つかりませんでした。');
      return null;
    }
    const csrfToken = match[1];
    Logger.log('情報: CSRFトークンを取得: ' + csrfToken);

    cookieMap = updateCookieMap(cookieMap, headers['Set-Cookie'] || headers['set-cookie']);
    Logger.log('情報: 初回セッションCookieを取得: ' + formatCookieHeader(cookieMap));
    
    addRandomDelay(500, 1000);

    const mandatoryParams = {
      'authenticity_token': csrfToken,
      'form[company_id]': credentials.companyId,
      'form[login_id]': credentials.loginId,
      'form[password]': credentials.password,
      'form[next]': '/ja',
      'commit': 'ログイン'
    };

    const repeatingParams = [
      ['form[fill_company_id_and_login_id]', '0'],
      ['form[fill_company_id_and_login_id]', '1']
    ];

    const allEntries = Object.entries(mandatoryParams).concat(repeatingParams);
    const payloadString = allEntries
      .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
      .join('&');

    const postOptions = {
      method: 'POST',
      payload: payloadString,
      contentType: 'application/x-www-form-urlencoded',
      followRedirects: false,
      muteHttpExceptions: true,
      headers: {
        'User-Agent': USER_AGENT,
        'Cookie': formatCookieHeader(cookieMap),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
        'Referer': LOGIN_URL,
        'Origin': 'https://atnd.ak4.jp',
        'Upgrade-Insecure-Requests': '1',
        'DNT': '1'
      }
    };

    const postResponse = UrlFetchApp.fetch(LOGIN_URL, postOptions);
    const statusCode = postResponse.getResponseCode();
    const postHeaders = postResponse.getAllHeaders();

    Logger.log(`情報: ログインPOST応答ステータス: ${statusCode}`);
    
    cookieMap = updateCookieMap(cookieMap, postHeaders['Set-Cookie'] || postHeaders['set-cookie']);

    if (statusCode === 302 || statusCode === 303) {
      const location = postHeaders['Location'] || postHeaders['location'];
      const locationUrl = location ? (location.startsWith('http') ? location : 'https://atnd.ak4.jp' + location) : '';
      Logger.log(`情報: POST成功ステータス ${statusCode} を受信。リダイレクト先: ${locationUrl || 'Locationヘッダーなし'}`);

      if (locationUrl && (locationUrl.includes(MANAGER_URL) || locationUrl === ROOT_JA_URL)) {
        Logger.log('情報: 最終セッションCookieを確定し、リダイレクト先のGETリクエストを実行します...');
        addRandomDelay(1500, 2500);
        
        const finalOptions = {
          method: 'GET',
          headers: {
            'User-Agent': USER_AGENT,
            'Cookie': formatCookieHeader(cookieMap),
            'Referer': LOGIN_URL,
            'Upgrade-Insecure-Requests': '1',
            'DNT': '1'
          },
          muteHttpExceptions: true,
          followRedirects: true
        };
        
        const finalResponse = UrlFetchApp.fetch(locationUrl, finalOptions);
        cookieMap = updateCookieMap(cookieMap, finalResponse.getAllHeaders()['Set-Cookie'] || finalResponse.getAllHeaders()['set-cookie']);
        
        const finalCookieString = formatCookieHeader(cookieMap);
        Logger.log('成功: ログイン成功。最終セッションCookie: ' + finalCookieString);
        addRandomDelay(3000, 5000);
        return finalCookieString;
      } else {
        Logger.log(`エラー: 302を受信しましたが、リダイレクト先が ${locationUrl} であったため、ログイン失敗（セッション確定失敗）と判断します。`);
        return null;
      }
    } 
    
    const postHtml = postResponse.getContentText('UTF-8');
    const toastMatch = postHtml.match(/<div class="c-toast\s*p-toast--runtime">([\s\S]*?)<\/div>/i);
    let errorMessage = 'エラーメッセージなし';
    if (toastMatch && toastMatch[1].trim() !== '') {
      errorMessage = toastMatch[1].trim().replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    }
    Logger.log(`エラー: ログイン失敗、ステータスコード: ${statusCode}, サーバー応答HTMLにエラー情報が含まれている可能性があります: ${errorMessage}`);
    return null;

  } catch (e) {
    Logger.log('エラー: ログイン処理に失敗: ' + e.message);
    return null;
  }
}

/**
 * 指定されたURLのHTMLを取得し、更新されたCookieマップを返す
 */
function fetchPageHTML(initialCookieMap, url) {
  let currentCookieMap = {...initialCookieMap};
  let retryCount = 0;
  const maxRetries = 3;

  while (retryCount < maxRetries) {
    addRandomDelay(2000, 4000);
    const options = {
      method: 'GET',
      headers: {
        'Cookie': formatCookieHeader(currentCookieMap),
        'User-Agent': USER_AGENT,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
        'Referer': getReferer(url),
        'Upgrade-Insecure-Requests': '1',
        'DNT': '1'
      },
      muteHttpExceptions: true,
      followRedirects: true
    };

    try {
      Logger.log(`情報: ${url} 取得試行 (${retryCount + 1}/${maxRetries})`);
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      const html = response.getContentText('UTF-8');
      Logger.log(`情報: ${url} ステータスコード: ${statusCode}`);

      currentCookieMap = updateCookieMap(currentCookieMap, response.getAllHeaders()['Set-Cookie'] || response.getAllHeaders()['set-cookie']);
      
      if (html.includes('<title>AKASHI - ログイン</title>')) {
        Logger.log('情報: ログインページにリダイレクト。セッション再取得を試みます...');
        const newCookieString = loginToAkashiWithRetry();
        if (!newCookieString) {
          Logger.log('エラー: 再ログイン失敗。処理を中断します。');
          return { html: '', cookieMap: currentCookieMap };
        }
        currentCookieMap = newCookieString.split('; ').reduce((acc, pair) => {
          const [key, value] = pair.split('=').map(s => s.trim());
          if (key) acc[key] = value;
          return acc;
        }, {});
        retryCount++;
        continue;
      }

      if (statusCode === 200) {
        return { html: html, cookieMap: currentCookieMap };
      } else {
        Logger.log(`エラー: ${url} 取得失敗: ステータスコード ${statusCode}`);
        retryCount++;
      }
    } catch (e) {
      Logger.log(`エラー: ${url} 取得中にエラー: ${e.message}`);
      retryCount++;
    }
  }

  Logger.log(`エラー: ${url} 取得で最大リトライ回数を超過。処理を中断します。`);
  return { html: '', cookieMap: currentCookieMap };
}

/**
 * URLに応じたRefererヘッダーを取得
 * @description ページ遷移のRefererヘッダーを決定します。
 * @param {string} url 対象URL
 * @returns {string} Refererヘッダーとして使用するURL
 */
function getReferer(url) {
  const refererMap = {
    [MANAGER_URL]: ROOT_JA_URL,
    [CURRENT_ATTENDANCE_URL]: MANAGER_URL,
    [ATTENDANCE_URL]: CURRENT_ATTENDANCE_URL
  };
  
  for (const key in refererMap) {
    if (url.startsWith(key)) return refererMap[key];
  }
  return LOGIN_URL;
}

/**
 * Google Chatにメッセージを送信
 * @param {string} message 送信するメッセージ
 * @param {boolean} includeMention メンション（<users/all>）を含むか
 */
function sendToGoogleChat(message, includeMention = true) {
  const credentials = getCredentials();
  const webhookUrl = credentials.webhookUrl;

  if (!webhookUrl) {
    Logger.log('エラー: WEBHOOK_URLが設定されていません。送信をスキップします。');
    return;
  }

  const payloadObject = includeMention ? {
    "text": "<users/all>" + message,
    "annotations": [{
      "type": "USER_MENTION",
      "start_index": 0,
      "length": 11,
      "user_mention": {
        "user": { "name": "users/all" }
      }
    }]
  } : {
    "text": message
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payloadObject),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const statusCode = response.getResponseCode();
    Logger.log(`情報: Google Chat送信ステータスコード: ${statusCode}`);
    if (statusCode !== 200) {
      Logger.log(`エラー: Google Chat送信エラー: ステータスコード ${statusCode}, レスポンス: ${response.getContentText()}`);
    }
  } catch (e) {
    Logger.log(`エラー: Google Chat送信エラー: ${e.message}`);
  }
}

/**
 * 勤怠HTMLを解析し、社員データを抽出
 */
function parseAttendanceHTML(html) {
  const employees = [];
  const rowRegex = /<tr id="daily_summary_\d+"[\s\S]*?<\/tr>/g;
  const rows = html.match(rowRegex) || [];
  Logger.log('情報: 抽出された行数: ' + rows.length);
  
  const parsedForLog = [];

  rows.forEach((row, index) => {
    let name = '';
    let scheduledStart = '--:--';
    let scheduledEnd = '--:--';
    let stampStart = '--:--';
    let stampEnd = '--:--';
    
    const cellContents = row.match(/<td[^>]*>([\s\S]*?)<\/td>/g) || [];
    if (cellContents.length < 14) return;

    try {
      let nameCell = cellContents[1];
      let cleanName = nameCell.replace(/<rt>[\s\S]*?<\/rt>/g, '');
      name = cleanName.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();

      let stampCell = cellContents[2];
      const stampTimeMatches = stampCell.match(/--:--|\d{2}:\d{2}/g) || [];
      stampStart = stampTimeMatches[0] || '--:--';
      stampEnd = stampTimeMatches[1] || '--:--';

      let workCell = cellContents[4];
      const scheduledTimeMatches = workCell.match(/(\d{2}:\d{2})/g) || [];
      scheduledStart = scheduledTimeMatches[0] || '--:--';
      scheduledEnd = scheduledTimeMatches[1] || '--:--';
      
      let reasonCell = cellContents[11];
      let reasonText = reasonCell.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
      const overtimeMatch = reasonText.match(/残業終了=(\d{2}:\d{2})/);
      if (overtimeMatch && overtimeMatch[1]) {
        const newEndTime = overtimeMatch[1];
        Logger.log(`情報: ${name} の理由欄から残業終了時刻 ${newEndTime} を検出しました。退勤予定を上書きします。`);
        scheduledEnd = newEndTime;
      }

      let statusCell = cellContents[5];
      let statusText = statusCell.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
      const leaveType = getLeaveType(statusText);
      
      if (leaveType) {
        switch (leaveType) {
          case '休暇':
            scheduledStart = '--:--';
            scheduledEnd = '--:--';
            Logger.log(`情報: ${name} は【${leaveType}】でしたので、出退勤予定を削除します。`);
            break;
          case '午前休':
            scheduledStart = '13:00';
            Logger.log(`情報: ${name} は【${leaveType}】でしたので、出勤予定を【13:00】に上書きします。`);
            break;
          case '午後休':
            scheduledEnd = '12:00';
            Logger.log(`情報: ${name} は【${leaveType}】でしたので、退勤予定を【12:00】に上書きします。`);
            break;
          default:
            Logger.log(`警告: ${name} の不明な休暇タイプ (${leaveType}) が検出されました。`);
            break;
        }
      }

      if (name && name.length > 0) {
        const employeeRecord = { 
          name, stampStart, stampEnd, scheduledStart, scheduledEnd
        };
        employees.push(employeeRecord);
        if (parsedForLog.length < 3) parsedForLog.push(employeeRecord);
      }
    } catch (e) {
      Logger.log(`エラー: 行 ${index} のパース中に例外が発生しました: ${e.message}`);
    }
  });
  
  return employees;
}

/**
 * ログイン機能をテスト
 * @description Akashiへのログインをテストし、セッションCookieの取得結果をログに記録します。
 */
function testLogin() {
  const sessionCookie = loginToAkashiWithRetry();
  if (sessionCookie) {
    Logger.log('情報: ログイン成功。セッション Cookie: ' + sessionCookie);
  } else {
    Logger.log('エラー: ログインに失敗しました。');
  }
}

/**
 * メイン処理をテスト
 * @description 勤怠チェックの全体処理をテスト実行します。
 */
function testMain() {
  main();
}
