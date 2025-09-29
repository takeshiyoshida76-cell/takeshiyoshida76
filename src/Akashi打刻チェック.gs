/**
 * Akashi勤怠チェック自動化スクリプト (Google Apps Script)
 * - 目的: Akashiの勤務表ページにアクセスし、出退勤の打刻漏れがないかチェックし、
 * Google Chat Webhook経由で通知を行います。
 * * * * 【仕様】:
 * 1. 勤怠サマリURLは日付なし (当日チェック) に固定。
 * 2. 出勤通知: 予定開始時刻の5分前以降、未打刻なら通知。
 * 3. 退勤通知: 予定終了時刻の15分後以降、未打刻なら通知。
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

// 勤怠サマリURL
const ATTENDANCE_URL = 'https://atnd.ak4.jp/ja/manager/daily_summary'; 

// User-Agent
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36';

// ==============================================================================
// 2. ユーティリティ関数
// ==============================================================================

function getCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    companyId: scriptProperties.getProperty('COMPANY_ID'),
    loginId: scriptProperties.getProperty('LOGIN_ID'),
    password: scriptProperties.getProperty('PASSWORD'),
    webhookUrl: scriptProperties.getProperty('WEBHOOK_URL')
  };
}

/**
 * 現在のJST日時を取得
 * @returns {{date: Date, dateString: string}} JST DateオブジェクトとYYYYMMDD形式の文字列
 */
function getCurrentJSTAndDateString() {
  const now = new Date();
  const jstTimeZone = 'Asia/Tokyo';
  
  // JSTでの現在時刻をDateオブジェクトとして取得
  const nowJST = new Date(Utilities.formatDate(now, jstTimeZone, 'yyyy-MM-dd HH:mm:ss') + '+09:00');
  
  // YYYYMMDD形式の文字列を生成 (ログ用)
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

function addRandomDelay(min, max) {
  const delay = Math.floor(Math.random() * (max - min + 1)) + min;
  Utilities.sleep(delay);
}

/**
 * Set-CookieヘッダーからCookieマップを更新する
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
 * CookieマップをHTTPリクエストヘッダー用の文字列に変換する
 */
function formatCookieHeader(cookieMap) {
  return Object.entries(cookieMap)
    .map(([key, value]) => `${key}=${value}`)
    .join('; ');
}


// ==============================================================================
// 3. メイン処理 (エントリーポイント)
// ==============================================================================

function main() {
  Logger.log('お知らせ: 実行開始');
  const sessionCookie = loginToAkashi();
  
  if (sessionCookie) {
    checkAttendance(sessionCookie);
  } else {
    Logger.log('エラー: ログインに失敗したため、処理を中断します。');
  }
  Logger.log('お知らせ: 実行完了');
}

function checkAttendance(sessionCookie) {
  const jstInfo = getCurrentJSTAndDateString();
  const now = jstInfo.date;
  
  Logger.log(`情報: 対象とする勤怠サマリURL (当日): ${ATTENDANCE_URL}`);
  
  const cookieMap = sessionCookie.split('; ').reduce((acc, pair) => {
    const [key, value] = pair.split('=').map(s => s.trim());
    if (key) acc[key] = value;
    return acc;
  }, {});
  
  // ページ遷移のシーケンス
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

  Logger.log('情報: 抽出された全従業員数: ' + employeeData.length);

  employeeData.forEach(employee => {
    // 氏名から名字のみを抽出 (例: "葭田 健志" -> "葭田")
    const lastName = employee.name.trim().split(/\s+/)[0]; 
    
    // TARGET_NAMESが空の場合は全員を対象にする。空でなければ名字がリストに含まれているかチェック。
    if (TARGET_NAMES.length === 0 || TARGET_NAMES.includes(lastName)) {
      
      const nowTime = now.getTime();

      Logger.log('チェック情報: 氏名:' + employee.name + ' 予定開始:' + employee.scheduledStart + ' 予定終了:' + employee.scheduledEnd + ' 打刻開始:' + employee.stampStart + ' 打刻終了:' + employee.stampEnd);
      // 1. 出勤打刻漏れチェック (予定開始時刻が設定されている場合)
      if (employee.scheduledStart !== '--:--') {
        const scheduledStartTime = parseTime(employee.scheduledStart, now);
        
        // 予定出勤時刻の5分前を通知開始時刻とする
        const fiveMinutesBefore = scheduledStartTime.getTime() - 5 * 60 * 1000;

        // 実績出勤時刻が未入力かつ、通知時刻（5分前）を過ぎている場合
        if (employee.stampStart === '--:--' && nowTime >= fiveMinutesBefore) {
          messages.push(`${employee.name} さん、予定出勤時刻 ${employee.scheduledStart} の出勤打刻が行われていません。至急、打刻してください。`);
        }
      }
      
      // 2. 退勤打刻漏れチェック (予定退勤時刻が設定されている場合)
      if (employee.scheduledEnd !== '--:--') {
        const scheduledEndTime = parseTime(employee.scheduledEnd, now);
        
        // 予定退勤時刻の15分後を通知開始時刻とする
        const fifteenMinutesAfter = scheduledEndTime.getTime() + 15 * 60 * 1000;
        
        // 実績退勤時刻が未入力かつ、通知時刻（15分後）を過ぎている場合
        if (employee.stampEnd === '--:--' && nowTime >= fifteenMinutesAfter) {
          messages.push(`${employee.name} さん、予定退勤時刻 ${employee.scheduledEnd} の退勤打刻が行われていません。至急、打刻してください。`);
        }
      }
    }
  });

  if (messages.length > 0) {
    const dateFormatted = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日');
    const header = `【打刻アラート】${dateFormatted} の打刻漏れが${messages.length}件あります！\n`;
    const messageText = header + messages.join('\n');
    sendToGoogleChat(messageText);
    Logger.log('情報: 送信したメッセージ:\n' + messageText);
  } else {
    Logger.log('情報: 送信するメッセージはありません。');
  }
}

// ==============================================================================
// 4. 外部通信関数 (Akashi & Google Chat)
// ==============================================================================

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
      .map(([k, v]) => 
        `${encodeURIComponent(k)}=${encodeURIComponent(v)}`
      )
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
        errorMessage = toastMatch[1].trim().replace(/<[^>]*>/g, '').replace(/\n/g, ' ').trim();
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
        const newCookieString = loginToAkashi(); 
        if (!newCookieString) {
          Logger.log('エラー: 再ログイン失敗。処理を中断します。');
          return { html: '', cookieMap: currentCookieMap };
        }
        const newCookieMap = newCookieString.split('; ').reduce((acc, pair) => {
          const [key, value] = pair.split('=').map(s => s.trim());
          if (key) acc[key] = value;
          return acc;
        }, {});
        currentCookieMap = {...newCookieMap}; 
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

function getReferer(url) {
  const refererMap = {
    [MANAGER_URL]: ROOT_JA_URL, 
    [CURRENT_ATTENDANCE_URL]: MANAGER_URL,
    [ATTENDANCE_URL]: CURRENT_ATTENDANCE_URL 
  };
  
  for (const key in refererMap) {
    if (url.startsWith(key)) {
      return refererMap[key];
    }
  }
  return LOGIN_URL;
}

function sendToGoogleChat(message) {
  const credentials = getCredentials();
  const webhookUrl = credentials.webhookUrl;

  if (!webhookUrl) {
    Logger.log('エラー: WEBHOOK_URLが設定されていません。送信をスキップします。');
    return;
  }

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
          }
        }
      }
    ]
  };

  // メッセージをJSON文字列に変換
  const payload = JSON.stringify(MESSAGES);
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const statusCode = response.getResponseCode();
    Logger.log(`情報: Google Chat送信ステータスコード: ${statusCode}`);
    if (statusCode !== 200) {
      Logger.log(`エラー: Google Chat送信エラー: ステータスコード ${statusCode}`);
    }
  } catch (e) {
    Logger.log(`エラー: Google Chat送信エラー: ${e.message}`);
  }
}

// ==============================================================================
// 5. データ解析関数
// ==============================================================================

function parseAttendanceHTML(html) {
  const employees = [];
  // <tr id="daily_summary_\d+" で始まる行を抽出
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
    
    // <td>...</td> の中身（コンテンツ部分）を全て抽出
    const cellContents = row.match(/<td[^>]*>([\s\S]*?)<\/td>/g) || [];
    
    // Akashiのテーブルは通常14列以上
    if (cellContents.length < 14) { 
        return; 
    }

    try {
        // --- 1. 氏名抽出 (Index 1) ---
        let nameCell = cellContents[1];
        let cleanName = nameCell.replace(/<rt>[\s\S]*?<\/rt>/g, '');
        name = cleanName.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();


        // --- 2. 実績打刻時刻の抽出 (Index 2) ---
        let stampCell = cellContents[2];
        const stampTimeMatches = stampCell.match(/--:--|\d{2}:\d{2}/g) || [];

        stampStart = stampTimeMatches[0] || '--:--';
        stampEnd = stampTimeMatches[1] || '--:--';


        // --- 3. 予定時刻の抽出 (Index 4) ---
        let workCell = cellContents[4];
        
        // 予定時刻の抽出
        const scheduledTimeMatches = workCell.match(/(\d{2}:\d{2})/g) || [];
        scheduledStart = scheduledTimeMatches[0] || '--:--';
        scheduledEnd = scheduledTimeMatches[1] || '--:--';
        
        if (name && name.length > 0) {
            const employeeRecord = { 
                name, 
                stampStart, 
                stampEnd, 
                scheduledStart, 
                scheduledEnd,
            };
            employees.push(employeeRecord);
            
            if (parsedForLog.length < 3) {
                parsedForLog.push(employeeRecord);
            }
        }
    } catch (e) {
        Logger.log(`エラー: 行 ${index} のパース中に例外が発生しました: ${e.message}`);
    }
  });
  
  return employees;
}

// ==============================================================================
// 6. テスト関数
// ==============================================================================

function testLogin() {
  const sessionCookie = loginToAkashi();
  if (sessionCookie) {
    Logger.log('情報: ログイン成功。セッション Cookie: ' + sessionCookie);
  } else {
    Logger.log('エラー: ログインに失敗しました。');
  }
}

function testMain() {
  main();
}
