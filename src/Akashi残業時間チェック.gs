// @ts-nocheck
/**
 * Akashi 月間残業15時間超アラート自動化スクリプト (Google Apps Script)
 *
 * 【概要】
 * - AKASHIの管理者ページから「月間残業時間」「月間所定休日労働時間」「月間法定休日労働時間」を取得し、
 * 合計が15時間以上の社員を検知します。
 * - 閾値超過者をGoogle Chatに@all通知します。
 * - 締め日までの残営業日数を計算し、通知メッセージに付加します。（CalendarAppを使用）
 *
 * 【使用方法】
 * 1. スクリプトプロパティに以下を設定: COMPANY_ID, LOGIN_ID, PASSWORD, CHAT_WEBHOOK_URL
 * 2. TARGET_EMPLOYEE_IDS を監視したい従業員IDの配列に変更
 * 3. トリガーで main() を毎日または毎月25日などに登録
 *
 * @fileoverview AKASHIの勤怠データを取得し、時間外労働が閾値を超えた場合に通知するGASスクリプト。
 * @requires CalendarApp (日本の祝日カレンダーへのアクセスに必要)
 */

// ==============================================================================
// 1. 定数設定
// ==============================================================================

/** @type {number[]} 監視対象従業員IDの配列 */
const TARGET_EMPLOYEE_IDS = [336546, 336415, 336530, 336540]; 
/** @type {number} 時間外労働の閾値（時間単位） */
const THRESHOLD_HOURS = 15;
/** @type {number} 時間外労働の閾値（分単位） */
const THRESHOLD_MINUTES = THRESHOLD_HOURS * 60;

/** @type {string} AKASHI ログインURL */
const LOGIN_URL = 'https://atnd.ak4.jp/ja/login';
/** @type {string} AKASHI 管理者ページURL */
const MANAGER_URL = 'https://atnd.ak4.jp/ja/manager';
/** @type {string} AKASHI ルート日本語URL */
const ROOT_JA_URL = 'https://atnd.ak4.jp/ja';
/** @type {string} HTTPリクエストに使用するUser-Agent */
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36';

// ==============================================================================
// 2. ユーティリティ関数
// ==============================================================================

/**
 * スクリプトプロパティから認証情報を取得します。
 * 必要なプロパティが設定されていない場合はエラーをスローします。
 * * @returns {{companyId: string, loginId: string, password: string}} 認証情報オブジェクト
 * @throws {Error} 認証情報が設定されていない場合
 */
function getCredentials() {
  const companyId = PropertiesService.getScriptProperties().getProperty('COMPANY_ID');
  const loginId = PropertiesService.getScriptProperties().getProperty('LOGIN_ID');
  const password = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  
  if (!companyId || !loginId || !password) {
    throw new Error('認証情報が設定されていません。');
  }
  return { companyId, loginId, password };
}

/**
 * 指定された範囲でランダムな遅延を発生させます。
 * * @param {number} minMs 最小遅延時間 (ミリ秒)
 * @param {number} maxMs 最大遅延時間 (ミリ秒)
 */
function addRandomDelay(minMs, maxMs) {
  const delay = Math.floor(Math.random() * (maxMs - minMs + 1)) + minMs;
  Utilities.sleep(delay);
}

/**
 * HTTPレスポンスの 'Set-Cookie' ヘッダーからCookieマップを更新します。
 * * @param {Object.<string, string>} cookieMap 現在のCookieマップ
 * @param {string|string[]|null|undefined} setCookieHeaders 'Set-Cookie' ヘッダーの値
 * @returns {Object.<string, string>} 更新されたCookieマップ
 */
function updateCookieMap(cookieMap, setCookieHeaders) {
  if (!setCookieHeaders) return cookieMap;
  if (!Array.isArray(setCookieHeaders)) setCookieHeaders = [setCookieHeaders];
  
  setCookieHeaders.forEach(header => {
    if (!header) return;
    const cookies = header.split(/,\s*(?=[^;]+=[^;])/); 
    cookies.forEach(cookie => {
      const [keyValue] = cookie.split(';');
      const [key, value] = keyValue.split('=', 2);
      if (key && value !== undefined) {
        cookieMap[key.trim()] = value.trim();
      }
    });
  });
  return cookieMap;
}

/**
 * CookieマップをHTTPヘッダー形式の文字列にフォーマットします。
 * * @param {Object.<string, string>} cookieMap Cookieマップ
 * @returns {string} フォーマットされたCookieヘッダー文字列
 */
function formatCookieHeader(cookieMap) {
  return Object.entries(cookieMap).map(([k, v]) => `${k}=${v}`).join('; ');
}

/**
 * Cookie文字列をCookieマップに変換します。
 * * @param {string} cookieStr Cookieヘッダー形式の文字列
 * @returns {Object.<string, string>} Cookieマップ
 */
function cookieStringToMap(cookieStr) {
  const map = {};
  if (cookieStr) {
    cookieStr.split('; ').forEach(pair => {
      const [k, v] = pair.split('=');
      if (k) map[k.trim()] = v ? v.trim() : '';
    });
  }
  return map;
}

/**
 * Google Chatにメッセージを送信します。
 * * @param {string} message 送信するメッセージ本文
 * @param {boolean} [includeMention=true] @all を含めるかどうか
 */
function sendToGoogleChat(message, includeMention = true) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
  if (!webhookUrl) return;
  
  const payload = { text: (includeMention ? '@all ' : '') + message };
  
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
  } catch (e) {
    Logger.log('エラー: Google Chatへの通知に失敗しました: ' + e.message);
  }
}

/**
 * 監視対象となる年と月をYYYYMM形式で取得します。
 * 22日以降は翌月を対象とします。
 * * @returns {string} YYYYMM形式の年月文字列
 */
function getTargetYearMonth() {
  const now = new Date();
  const day = now.getDate();
  let year = now.getFullYear();
  let month = now.getMonth() + 1; // getMonth()は0-11

  // 22日以降は翌月を対象とする
  if (day >= 22) {
    month += 1;
    if (month > 12) { 
      month = 1; 
      year += 1; 
    }
  }
  
  return `${year}${month.toString().padStart(2, '0')}`;
}

/**
 * HH:MM形式の時間を分単位に変換します。
 * * @param {string} str HH:MM形式の時刻文字列 (例: "12:30", "100:05")
 * @returns {number} 分単位に変換された時間、または無効な場合は0
 */
function timeToMinutes(str) {
  if (!str || !/^\d+:\d{2}$/.test(str)) return 0;
  const [h, m] = str.split(':').map(Number);
  return h * 60 + m;
}

/**
 * 分単位の時間をHH:MM形式に変換します。
 * * @param {number} mins 分単位の時間
 * @returns {string} HH:MM形式の時刻文字列 (例: "15:00", "0:30")
 */
function minutesToTime(mins) {
  if (mins < 0) mins = 0;
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h}:${m.toString().padStart(2, '0')}`;
}

/**
 * HTMLからCSRFトークンを抽出します。
 * * @param {string} html ページのHTMLコンテンツ
 * @returns {string|null} CSRFトークン、または見つからない場合はnull
 */
function extractCsrfToken(html) {
  const match = html.match(/<meta name="csrf-token" content="([^"]+)" \/>/);
  return match ? match[1] : null;
}

/**
 * 指定されたDateオブジェクトを YYYY/MM/DD 形式の文字列にフォーマットします。
 * * @param {Date} date フォーマット対象のDateオブジェクト
 * @returns {string} YYYY/MM/DD形式の文字列
 */
function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

/**
 * 指定されたURLからHTMLコンテンツをフェッチします。
 * ログインセッションの維持とリトライロジックを含みます。
 * * @param {Object.<string, string>} initialCookieMap 初期Cookieマップ
 * @param {string} url フェッチするURL
 * @returns {{html: string, cookieMap: Object.<string, string>}} 取得したHTMLと更新されたCookieマップ
 */
function fetchPageHTML(initialCookieMap, url) {
  let cookieMap = { ...initialCookieMap };
  let retryCount = 0;
  const maxRetries = 3;

  while (retryCount < maxRetries) {
    addRandomDelay(2000, 4000);
    const options = {
      method: 'GET',
      headers: {
        'Cookie': formatCookieHeader(cookieMap),
        'User-Agent': USER_AGENT,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Referer': url.includes('attendance') ? MANAGER_URL : LOGIN_URL,
      },
      muteHttpExceptions: true,
      followRedirects: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const html = response.getContentText('UTF-8');
      cookieMap = updateCookieMap(cookieMap, response.getAllHeaders()['Set-Cookie']);

      // ログイン切れを検知した場合、再ログインを試みる
      if (html.includes('<title>AKASHI - ログイン</title>')) {
        Logger.log('情報: セッション切れのため、再ログインを試みます。');
        const newCookieString = loginToAkashiWithRetry();
        if (newCookieString) {
          cookieMap = cookieStringToMap(newCookieString);
        } else {
          // 再ログイン失敗
          Logger.log('エラー: 再ログインに失敗しました。処理を中断します。');
          return { html: '', cookieMap: {} }; 
        }
        retryCount++; 
        continue;
      }
      
      if (response.getResponseCode() === 200) return { html, cookieMap };

    } catch (e) {
      Logger.log(`警告: URLフェッチ中に例外が発生しました (${url}): ${e.message}`);
      retryCount++;
    }
  }
  Logger.log(`エラー: ${url} のフェッチが最大リトライ回数 (${maxRetries}) を超えました。`);
  return { html: '', cookieMap };
}

/**
 * AKASHIログイン処理を実行します。
 * * @return {string|null} 最終Cookie文字列、またはログイン失敗時はnull
 */
function loginToAkashi() {
  Logger.log('情報: ログイン処理開始');
  const credentials = getCredentials();
  
  let cookieMap = {};
  
  // 1. ログインページにアクセスし、CSRFトークンを取得
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
    
    // CSRFトークンを抽出
    const match = html.match(/<input[^>]*name="authenticity_token"[^>]*value="([^"]*)"/);
    if (!match || !match[1]) {
      Logger.log('エラー: ログインページからCSRFトークンが見つかりませんでした。');
      return null;
    }
    const csrfToken = match[1];

    cookieMap = updateCookieMap(cookieMap, headers['Set-Cookie'] || headers['set-cookie']);

    addRandomDelay(500, 1000); // 遅延

    // 2. ログイン情報をPOST送信
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
    const payloadString = allEntries.map(([k, v]) => 
      `${encodeURIComponent(k)}=${encodeURIComponent(v)}`
    ).join('&');

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

    cookieMap = updateCookieMap(cookieMap, postHeaders['Set-Cookie'] || postHeaders['set-cookie']);

    // 3. リダイレクト先の最終ページをフェッチ（ログイン成功時）
    if (statusCode === 302 || statusCode === 303) {
      const location = postHeaders['Location'] || postHeaders['location'];
      const locationUrl = location ? 
        (location.startsWith('http') ? location : 'https://atnd.ak4.jp' + location) : 
        '';

      if (locationUrl && (locationUrl.includes(MANAGER_URL) || locationUrl === ROOT_JA_URL)) {
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
        Logger.log('成功: ログイン成功。');
        addRandomDelay(3000, 5000); 
        return finalCookieString;
      } else {
        Logger.log(`エラー: リダイレクト先が不正です: ${locationUrl}`);
        return null;
      }
    }

    // 4. ステータスコードが200の場合、エラーメッセージを抽出
    const postHtml = postResponse.getContentText('UTF-8');
    const toastMatch = postHtml.match(/<div class="c-toast\s*p-toast--runtime">([\s\S]*?)<\/div>/i);
    let errorMessage = 'エラーメッセージなし';
    if (toastMatch && toastMatch[1].trim() !== '') {
      errorMessage = toastMatch[1].trim().replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    }
    Logger.log(`エラー: ログイン失敗: ステータスコード ${statusCode}, メッセージ: ${errorMessage}`);
    return null;

  } catch (e) {
    Logger.log('エラー: ログイン処理で例外が発生しました: ' + e.message);
    return null;
  }
}

/**
 * AKASHIログイン処理を複数回リトライするラッパー関数。
 * * @param {number} [maxRetries=3] 最大リトライ回数
 * @return {string|null} ログイン成功時のCookie文字列、または全リトライ失敗時はnull
 */
function loginToAkashiWithRetry(maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    const cookie = loginToAkashi();
    if (cookie) return cookie;
    addRandomDelay(5000, 10000);
  }
  return null;
}

// ==============================================================================
// 3. 月間合計取得専用関数
// ==============================================================================

/**
 * AKASHIから指定された従業員の月間サマリーデータをフェッチします。
 * * @param {Object.<string, string>} cookieMap 現在のCookieマップ
 * @param {number} empId 従業員ID
 * @param {string} yearMonth YYYYMM形式の年月文字列
 * @returns {string} サマリーデータのレスポンステキスト (HTML断片/JS)
 */
function fetchSummaryData(cookieMap, empId, yearMonth) {
  const url = `https://atnd.ak4.jp/ja/manager/attendance/summary/${empId}/${yearMonth}?_=${Date.now()}`;
  
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Cookie': formatCookieHeader(cookieMap),
      'User-Agent': USER_AGENT,
      'Accept': 'text/javascript,*/*;q=0.01',
      'X-Requested-With': 'XMLHttpRequest',
      'Referer': `https://atnd.ak4.jp/ja/manager/attendance/${empId}/${yearMonth}`
    },
    muteHttpExceptions: true
  });
  
  updateCookieMap(cookieMap, res.getAllHeaders()['Set-Cookie']);
  
  return res.getContentText();
}

/**
 * サマリーデータテキストから月間残業時間、所定休日労働時間、法定休日労働時間の合計（分単位）をパースします。
 * * @param {string} text AKASHIから取得したサマリーデータテキスト（HTML断片）
 * @returns {number} 合計時間（分単位）
 */
function parseTotalOvertimeMinutes(text) {
  let total = 0;

  // 月間残業時間、所定休日労働時間、法定休日労働時間を抽出する正規表現
  const regs = [
    /月間残業時間.*?p-summary-roster-table-body__value[^>]*>(?:\\n|[\s\S])*?(\d+):(\d{2})/m,
    /月間所定休日労働時間.*?p-summary-roster-table-body__value[^>]*>(?:\\n|[\s\S])*?(\d+):(\d{2})/m,
    /月間法定休日労働時間.*?p-summary-roster-table-body__value[^>]*>(?:\\n|[\s\S])*?(\d+):(\d{2})/m
  ];

  regs.forEach((re) => {
    const m = text.match(re);

    if (m) {
      const h = parseInt(m[1], 10);
      const mm = parseInt(m[2], 10);
      total += h * 60 + mm;
      // 元のコードでコメントアウトされていたLogger.logはそのまま維持します
      // const names = ['残業', '所定休日', '法定休日']; 
      // Logger.log(`検出 → 月間${names[i]}労働時間: ${h}:${mm.toString().padStart(2,'0')}`);
    }
  });
  return total;
}

// ==============================================================================
// 4. 残営業日取得関数 (CalendarApp利用)
// ==============================================================================

/**
 * 締め日（前日の日付け基準で判断される締め日、例: 20日）のDateオブジェクトを取得します。
 * * @returns {Date} 締め日のDateオブジェクト
 */
function getClosingDate() {
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1); 

  const year = yesterday.getFullYear();
  const month = yesterday.getMonth(); // 0-11
  const day = yesterday.getDate();

  let closingYear, closingMonth;
  // 前日が20日以下の場合、当月20日が締め日
  if (day <= 20) {
    closingYear = year;
    closingMonth = month;
  } 
  // 前日が21日以上の場合、翌月20日が締め日
  else {
    closingMonth = month + 1;
    closingYear = year;
    if (closingMonth > 11) { // 12月の場合の年越し処理
      closingMonth = 0; // 1月
      closingYear++;
    }
  }
  
  // 締め日の20日を設定
  return new Date(closingYear, closingMonth, 20);
}

/**
 * 前日から締め日までの営業日数を計算します。
 * Google公式の「日本の祝日」カレンダーを参照し、土日祝日を除外します。
 * * @param {Date} closingDate 締め日のDateオブジェクト (getClosingDateで取得)
 * @returns {number} 前日から締め日までの営業日数
 * @throws {Error} 日本の祝日カレンダーへのアクセスに失敗した場合（CalendarApp権限不足など）
 */
function countBusinessDaysToClosing(closingDate) {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);

  // Google公式「日本の祝日」カレンダーから祝日を取得
  const holidayCalendar = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");
  if (!holidayCalendar) {
    throw new Error("日本の祝日カレンダーにアクセスできません。スクリプトにカレンダー権限を付与してください。");
  }

  // 検索期間を yesterday から closingDate の翌日までとする
  const searchStart = new Date(yesterday);
  const searchEnd = new Date(closingDate);
  searchEnd.setDate(searchEnd.getDate() + 1);

  const holidayEvents = holidayCalendar.getEvents(searchStart, searchEnd);
  // 祝日リストを Set に変換 ("YYYY-MM-DD"形式)
  const holidaySet = new Set(holidayEvents
    .filter(event => event.isAllDayEvent()) 
    .map(event => {
      const date = event.getStartTime();
      date.setHours(0, 0, 0, 0); 
      return date.toISOString().slice(0, 10); 
    })
  ); 

  let current = new Date(yesterday);
  current.setHours(0, 0, 0, 0);
  let count = 0;

  // 前日から締め日までをループ
  while (current <= closingDate) {
    const dateStr = current.toISOString().slice(0, 10);
    const dayOfWeek = current.getDay(); // 日曜日=0, 土曜日=6

    // 土日 (0, 6) ではなく、祝日でもない → 営業日
    if (dayOfWeek !== 0 && dayOfWeek !== 6 && !holidaySet.has(dateStr)) {
      count++;
    }

    current.setDate(current.getDate() + 1);
  }

  return count;
}


// ==============================================================================
// メイン関数（これだけトリガーに登録）
// ==============================================================================

/**
 * メイン処理を実行します。
 * 監視対象従業員の月間時間外労働時間を取得し、閾値を超過した場合はGoogle Chatに通知します。
 * * @returns {void}
 */
function main() {
  const yearMonth = getTargetYearMonth();
  const disp = `${yearMonth.slice(0, 4)}年${parseInt(yearMonth.slice(4))}月`;
  Logger.log(`【時間外労働${THRESHOLD_HOURS}時間以上チェック開始】${disp}`);

  // ログインとリトライ
  const cookieString = loginToAkashiWithRetry();
  if (!cookieString) {
    sendToGoogleChat(`【失敗】${disp} ログインできませんでした`, false);
    return;
  }

  let cookieMap = cookieStringToMap(cookieString);
  const alerts = [];
  const results = [];

  // 従業員ごとにループしてデータを取得・チェック
  for (const empId of TARGET_EMPLOYEE_IDS) {
    addRandomDelay(3000, 7000);

    // 1. 社員名を取得するための勤怠ページをフェッチ
    const page = fetchPageHTML(cookieMap, `https://atnd.ak4.jp/ja/manager/attendance/${empId}/${yearMonth}`);
    
    if (!page.html || !page.html.includes('p-roster-header__name')) {
      Logger.log(`警告: 従業員ID ${empId} のページ取得に失敗したか、社員名が見つかりませんでした。スキップします。`);
      continue; 
    }
    cookieMap = page.cookieMap;

    // 社員名をHTMLからパース
    const nameRaw = page.html.match(/<h2 class="p-roster-header__name">([\s\S]*?)<\/h2>/)?.[1] || '';
    const name = nameRaw
      .replace(/<rt>[\s\S]*?<\/rt>/gi, '') // ふりがなタグを除去
      .replace(/<[^>]+>/g, '') // その他のHTMLタグを除去
      .replace(/\s+/g, ' ')
      .trim();

    // 2. 月間サマリーデータをXHRでフェッチ
    const summaryText = fetchSummaryData(cookieMap, empId, yearMonth);
    if (!summaryText) {
      results.push({ name, time: '取得失敗' });
      Logger.log(`${name} → サマリーデータ取得失敗`);
      continue;
    }

    // 3. 時間外労働時間を解析し、合計時間を算出
    const mins = parseTotalOvertimeMinutes(summaryText);
    const timeStr = minutesToTime(mins);
    Logger.log(`${name} → 月間時間外労働時間: ${timeStr}`);
    results.push({ name, time: timeStr });

    // 4. 閾値チェック
    if (mins >= THRESHOLD_MINUTES) {
      alerts.push({ name, time: timeStr });
    }
  }

  // 締め日と残営業日数を取得
  let closingDate, businessDays;
  try {
    closingDate = getClosingDate();
    businessDays = countBusinessDaysToClosing(closingDate);
    Logger.log(`情報: 締め日: ${formatDate(closingDate)} までの残営業日日数: ${businessDays}`);
  } catch (e) {
    // CalendarApp権限エラーなどが発生した場合
    Logger.log(`エラー: 残営業日数の計算中にエラーが発生しました: ${e.message}`);
    businessDays = '計算エラー（権限確認）'; // エラーメッセージを変更
  }

  // 最終結果をGoogle Chatに通知
  if (alerts.length > 0) {
    let msg = `【月間時間外労働${THRESHOLD_HOURS}時間以上アラート】${disp}\n\n`;
    
    // 超過者リスト
    alerts.forEach(a => msg += `${a.name} さん → ${a.time}\n`);
    
    // 補足情報
    msg += `\n※残業＋所定休日労働＋法定休日労働の合計です`;
    msg += `\n\n【残営業日情報】`;
    msg += `\n締め日 ${formatDate(closingDate)} までの残営業日数は ${businessDays} 日です。`;
    msg += `\n未申請で閾値超過の可能性があれば、速やかに申請してください。`;
    
    sendToGoogleChat(msg, true);
    Logger.log('情報: Chat通知送信完了');
  } else {
    Logger.log('情報: 該当者なし → 通知なし');
  }

  Logger.log('【チェック完了】');
}
