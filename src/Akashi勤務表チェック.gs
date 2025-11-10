/**
 * Akashi勤務実績チェック自動化スクリプト (Google Apps Script)
 * 
 * 【概要】
 * - AKASHIの管理者ページから従業員の月次勤務データを取得し、勤怠ルール違反をチェック。
 * - Google Chatに指摘事項を通知。
 * - フレックス/ノーマル労働制対応、未来日スキップ、理由抽出強化。
 * 
 * 【使用方法】
 * 1. スクリプトプロパティ設定: COMPANY_ID, LOGIN_ID, PASSWORD, CHAT_WEBHOOK_URL
 * 2. TARGET_EMPLOYEE_IDS に従業員ID配列設定。
 * 3. トリガー設定: 毎日実行推奨 (22日以降で翌月チェック)。
 */

// ==============================================================================
// 1. 定数設定
// ==============================================================================

const TARGET_EMPLOYEE_IDS = [336546, 336415, 336530, 336540]; // 対象従業員番号 (複数可)
const LOGIN_URL = 'https://atnd.ak4.jp/ja/login';
const MANAGER_URL = 'https://atnd.ak4.jp/ja/manager';
const ROOT_JA_URL = 'https://atnd.ak4.jp/ja'; 
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36';

// ==============================================================================
// 2. ユーティリティ関数 (元スクリプト流用)
// ==============================================================================
/**
 * スクリプトプロパティから認証情報を取得
 * @return {Object} { companyId, loginId, password }
 */
function getCredentials() {
  const companyId = PropertiesService.getScriptProperties().getProperty('COMPANY_ID');
  const loginId = PropertiesService.getScriptProperties().getProperty('LOGIN_ID');
  const password = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (!companyId || !loginId || !password) throw new Error('認証情報が設定されていません。');
  return { companyId, loginId, password };
}

/**
 * JST現在日時とyyyyMMdd文字列を取得
 * @return {Object} { date: Date, dateStr: string }
 */
function getCurrentJSTAndDateString() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, "JST", "yyyyMMdd");
  const date = new Date(now); // JST に変換せずにそのまま使用（必要に応じて調整）
  return { date, dateStr };
}

/**
 * ランダムディレイ (ms)
 * @param {number} minMs 最小遅延
 * @param {number} maxMs 最大遅延
 */
function addRandomDelay(minMs, maxMs) {
  const delay = Math.floor(Math.random() * (maxMs - minMs + 1)) + minMs;
  Utilities.sleep(delay);
}

/**
 * Cookieマップ更新
 * @param {Object} cookieMap 既存マップ
 * @param {string|Array} setCookieHeaders Set-Cookieヘッダー
 * @return {Object} 更新後マップ
 */
function updateCookieMap(cookieMap, setCookieHeaders) {
  if (!setCookieHeaders) return cookieMap;
  if (!Array.isArray(setCookieHeaders)) setCookieHeaders = [setCookieHeaders];
  setCookieHeaders.forEach(header => {
    if (!header) return;
    const cookies = header.split(/,\s*(?=[^;]+=[^;])/);
    cookies.forEach(cookie => {
      const [keyValue, ...attrs] = cookie.split(';');
      const [key, value] = keyValue.split('=', 2);
      if (key && value !== undefined) cookieMap[key.trim()] = value.trim();
    });
  });
  return cookieMap;
}

/**
 * Cookieマップをヘッダー文字列に変換
 * @param {Object} cookieMap
 * @return {string}
 */
function formatCookieHeader(cookieMap) {
  return Object.entries(cookieMap).map(([k, v]) => `${k}=${v}`).join('; ');
}

/**
 * Refererヘッダー取得 (簡易)
 * @param {string} url
 * @return {string}
 */
function getReferer(url) {
  if (url.includes('attendance')) return MANAGER_URL;
  return LOGIN_URL;
}

/**
 * Google Chatにメッセージ送信
 * @param {string} message
 * @param {boolean} includeMention @all付与
 */
function sendToGoogleChat(message, includeMention = true) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
  if (!webhookUrl) return;
  const payload = { text: (includeMention ? '@all ' : '') + message };
  UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}

/**
 * ログインリトライラッパー
 * @param {number} maxRetries
 * @return {string|null} Cookie文字列
 */
function loginToAkashiWithRetry(maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    const cookie = loginToAkashi();
    if (cookie) return cookie;
    addRandomDelay(5000, 10000);
  }
  return null;
}

/**
 * AKASHIログイン処理
 * @return {string|null} 最終Cookie文字列
 */
function loginToAkashi() {
  Logger.log('お知らせ: ログイン処理開始');
  const credentials = getCredentials();
  if (!credentials.companyId || !credentials.loginId || !credentials.password) {
    Logger.log('エラー: 認証情報が不足しています。');
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

    cookieMap = updateCookieMap(cookieMap, headers['Set-Cookie'] || headers['set-cookie']);
    
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
    const payloadString = allEntries.map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`).join('&');

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

    if (statusCode === 302 || statusCode === 303) {
      const location = postHeaders['Location'] || postHeaders['location'];
      const locationUrl = location ? (location.startsWith('http') ? location : 'https://atnd.ak4.jp' + location) : '';
      
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
        Logger.log(`エラー: リダイレクト先不正: ${locationUrl}`);
        return null;
      }
    } 
    
    const postHtml = postResponse.getContentText('UTF-8');
    const toastMatch = postHtml.match(/<div class="c-toast\s*p-toast--runtime">([\s\S]*?)<\/div>/i);
    let errorMessage = 'エラーメッセージなし';
    if (toastMatch && toastMatch[1].trim() !== '') {
      errorMessage = toastMatch[1].trim().replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    }
    Logger.log(`エラー: ログイン失敗: ${statusCode}, ${errorMessage}`);
    return null;

  } catch (e) {
    Logger.log('エラー: ログイン例外: ' + e.message);
    return null;
  }
}

/**
 * HTMLページ取得 (Cookie維持、再ログイン対応)
 * @param {Object} initialCookieMap
 * @param {string} url
 * @return {Object} { html: string, cookieMap: Object }
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
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      const html = response.getContentText('UTF-8');

      currentCookieMap = updateCookieMap(currentCookieMap, response.getAllHeaders()['Set-Cookie'] || response.getAllHeaders()['set-cookie']);
      
      if (html.includes('<title>AKASHI - ログイン</title>')) {
        const newCookieString = loginToAkashiWithRetry();
        if (!newCookieString) return { html: '', cookieMap: currentCookieMap };
        currentCookieMap = cookieStringToMap(newCookieString);
        retryCount++;
        continue;
      }

      if (statusCode === 200) {
        return { html, cookieMap: currentCookieMap };
      } else {
        retryCount++;
      }
    } catch (e) {
      retryCount++;
    }
  }

  return { html: '', cookieMap: currentCookieMap };
}

/**
 * 時間文字列判定
 * @param {string} str
 * @return {boolean}
 */
function isTime(str) {
  if (!str || str === '--:--' || str === '') return false;
  const parts = str.split(':');
  if (parts.length !== 2) return false;
  const h = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  return !isNaN(h) && !isNaN(m) && h >= 0 && h < 24 && m >= 0 && m < 60;
}

/**
 * 時間を分に変換
 * @param {string} str
 * @return {number}
 */
function timeToMinutes(str) {
  if (!isTime(str)) return 0;
  const [h, m] = str.split(':').map(Number);
  return h * 60 + m;
}

/**
 * 分を時間文字列に変換
 * @param {number} mins
 * @return {string}
 */
function minutesToTime(mins) {
  if (mins < 0) mins = 0;
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
}

/**
 * 理由キーワードチェック
 * @param {string} str
 * @return {boolean}
 */
function checkReason(str) {
  if (!str) return false;
  const keywords = ['理由', 'ため', '為', 'に基づく', '会議', 'による'];
  return keywords.some(k => str.includes(k));
}

/**
 * CSRFトークン抽出
 * @param {string} html
 * @return {string|null}
 */
function extractCsrfToken(html) {
  const match = html.match(/<meta name="csrf-token" content="([^"]+)" \/>/);
  return match ? match[1] : null;
}

/**
 * 勤怠記録を取得
 * @param {Map} cookieMap
 * @param {number} empId
 * @param {string} yearMonth
 * @param {string} csrfToken
 * @return {string|null}
 */
function fetchAttendanceRecords(cookieMap, empId, yearMonth, csrfToken) {
  const timestamp = Date.now();
  const url = `https://atnd.ak4.jp/ja/manager/attendance/records/${empId}/${yearMonth}/?_=${timestamp}`;
  const headers = {
    'User-Agent': USER_AGENT,
    'Accept': 'text/javascript, application/javascript, application/ecmascript, application/x-ecmascript, */*; q=0.01',
    'X-CSRF-Token': csrfToken,
    'X-Requested-With': 'XMLHttpRequest',
    'Referer': `https://atnd.ak4.jp/ja/manager/attendance/${empId}/${yearMonth}`,
    'Cookie': formatCookieHeader(cookieMap),
    'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
    'DNT': '1'
  };
  const options = { method: 'get', headers, muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: false };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code !== 200) {
    Logger.log(`XHR失敗 Code: ${code}`);
    return null;
  }
  updateCookieMap(cookieMap, response.getHeaders()['Set-Cookie'] || []);
  return text;
}

/**
 * 対象年月取得 (22日以降で翌月)
 * @return {string} yyyyMM
 */
function getTargetYearMonth() {
  const { date: nowJST } = getCurrentJSTAndDateString();
  const day = nowJST.getDate();
  let year = nowJST.getFullYear();
  let month = nowJST.getMonth() + 1;
  if (day >= 22) {
    month += 1;
    if (month > 12) { month = 1; year += 1; }
  }
  return `${year}${month.toString().padStart(2, '0')}`;
}

/**
 * Cookie文字列をマップに変換
 * @param {string} cookieStr
 * @return {Object}
 */
function cookieStringToMap(cookieStr) {
  const map = {};
  if (cookieStr) {
    cookieStr.split('; ').forEach(pair => {
      const [k, v] = pair.split('=');
      if (k) map[k.trim()] = v.trim();
    });
  }
  return map;
}

/**
 * JSから記録解析
 * @param {string} jsText
 * @param {string} pageHtml
 * @param {number} empId
 * @return {Object|null} { name, formatType, days[] }
 */
function parseRecordsFromJs(jsText, pageHtml, empId) {
  // 名前抽出: ruby/rt除去、漢字のみ
  const nameMatch = pageHtml.match(/<h2 class="p-roster-header__name">([\s\S]*?)<\/h2>/);
  let name = `従業員${empId}`;
  if (nameMatch) {
    const rubyText = nameMatch[1].replace(/<rt>[\s\S]*?<\/rt>/g, '').replace(/<\/?ruby>/g, '');
    name = rubyText.replace(/\s+/g, '').trim();
  }
  Logger.log(name + 'さんのチェック');

  // formatType & recordsId
  let recordsId = '', formatType = '';
  const flexMatch = jsText.match(/data-attendance-records="flex_(\d+)"/);
  const normalMatch = jsText.match(/data-attendance-records="normal_labor_(\d+)"/);
  if (flexMatch) {
    recordsId = `flex_${flexMatch[1]}`; formatType = 'flex';
  } else if (normalMatch) {
    recordsId = `normal_labor_${normalMatch[1]}`; formatType = 'normal';
  } else {
    return null;
  }

  // replaceWith 内の HTML 抽出
  let tableHtml = '';
  const htmlMatch = jsText.match(/replaceWith\s*\(\s*(["'])((?:\\[\s\S]|[^\\])*?)\1\s*\)/);
  if (htmlMatch && htmlMatch[2]) {
    let rawHtml = htmlMatch[2]
      .replace(/\\"/g, '"')
      .replace(/\\\//g, '/')
      .replace(/\\t/g, '\t')
      .replace(/\\r/g, '')
      .replace(/\\n/g, '\n')
      .replace(/=\\"/g, '="')
      .replace(/\\">/g, '">');
    const tableMatch = rawHtml.match(/<table\s+class=(?:["']c-main-table-none-sidespace p-roster-main-table polyfill-sticky["']|\\"c-main-table-none-sidespace p-roster-main-table polyfill-sticky\\")[^>]*>([\s\S]*<\/table>)/);
    if (tableMatch && tableMatch[0]) {
      tableHtml = tableMatch[0];
    } else {
      Logger.log(`<table> 抽出失敗 from jsText: ${rawHtml.substring(0, 6000)}`);
    }
  } else {
    // フォールバック: jsText 全体から <table> を検索
    const tableMatch = jsText.match(/<table\s+class=(?:["']c-main-table-none-sidespace p-roster-main-table polyfill-sticky["']|\\"c-main-table-none-sidespace p-roster-main-table polyfill-sticky\\")[^>]*>([\s\S]*<\/table>)/);
    if (tableMatch && tableMatch[0]) {
      tableHtml = tableMatch[0];
    } else {
      // 最終手段: pageHtml からテーブルを抽出
      const pageTableMatch = pageHtml.match(/<table[^>]*class=["']c-main-table-none-sidespace p-roster-main-table polyfill-sticky["'][^>]*>[\s\S]*?<\/table>/);
      if (pageTableMatch) {
        tableHtml = pageTableMatch[0];
      } else {
        Logger.log(`pageHtmlからテーブル抽出失敗: ${pageHtml.substring(0, 1000)}`);
        return null;
      }
    }
  }

  const days = [];
  // <tr> 要素を抽出（各行を個別にキャプチャし、理由行も含む）
  const rowMatches = tableHtml.match(/<tr id="working_report_\d{8}"[\s\S]*?(?:<tr class="c-main-table-body__row--\w+"[\s\S]*?<\/tr>|\s*(?=<tr id="working_report_\d{8}"|\s*<\/tbody>))/g) || [];

  if (!rowMatches.length) {
    Logger.log(`rowMatches が見つかりません: ${tableHtml.substring(0, 6000)}`);
  }

  rowMatches.forEach(block => {
    const dateMatch = block.match(/working_report_(\d{8})/);
    if (!dateMatch) {
      Logger.log(`日付マッチ失敗: ${block.substring(0, 500)}`);
      return;
    }
    const dateStr = dateMatch[1];

    // 理由抽出: second-line-fields 内の理由を対象
    let reason = '';
    const reasonBlockMatch = block.match(/<tr class="c-main-table-body__row--\w+"[\s\S]*?<td[^>]*class="c-main-table-body__cell second-line-fields"[^>]*>[\s\S]*?理由[\s\S]*?<\/td>[\s\S]*?<\/tr>/);
    if (reasonBlockMatch) {
      let tdContent = reasonBlockMatch[0].match(/<td[^>]*class="c-main-table-body__cell second-line-fields"[^>]*>([\s\S]*?)<\/td>/)[1].trim();
      //reason = tdContent.replace(/<span[^>]*>理由<\/span>/i, '').trim() || '';
      const reasonRegex = /理由\n\s*?<\/span>[\s\S]*?([^<]+)[\s\S]*?<\/td>/;
      const reasonMatch = block.match(reasonRegex);
      if (reasonMatch && reasonMatch.length > 1) {
       reason = reasonMatch[1].trim(); 
      } else {
         reason = '';
      }
    } else {
       reason = '';
    }

    const mainTrMatch = block.match(/<tr id="working_report_.*?>([\s\S]*?)<\/tr>/);
    if (!mainTrMatch) {
      Logger.log(`メインTRマッチ失敗 (${dateStr}): ${block.substring(0, 500)}`);
      return;
    }
    const cells = mainTrMatch[1].match(/<td[^>]*>([\s\S]*?)<\/td>/g) || [];
    if (cells.length < (formatType === 'flex' ? 9 : 10)) {
      Logger.log(`セル数不足 (${dateStr}): ${cells.length} 個見つかりました`);
      return;
    }

    const getCellText = (index) => cells[index] ? cells[index].replace(/<td[^>]*>/, '').replace(/<\/td>/g, '').trim().replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() : '';
    const getTimesFromCell = (index, lines) => {
      const cellHtml = cells[index] || '';
      const spans = cellHtml.match(/<span[^>]*>([^<]+)<\/span>/g) || [];
      const times = spans.map(s => s.match(/<span[^>]*>([^<]+)<\/span>/)[1].trim());
      return times.length >= lines ? times.slice(0, lines) : times.concat(Array(lines - times.length).fill('--:--'));
    };

    const clockTimes = getTimesFromCell(1, 2);
    const actualTimes = getTimesFromCell(2, 2);
    const plannedTimes = getTimesFromCell(3, 2);
    const statusCell = getCellText(4);
    const totalTime = getCellText(5);
    const breakTime = getCellText(6);
    const nightTime = getCellText(7);
    let overtime = formatType === 'normal' ? getCellText(8) : (isTime(totalTime) ? minutesToTime(timeToMinutes(totalTime) - 480) : '0:00');
    const lateTime = getCellText(formatType === 'normal' ? 9 : 8);

    days.push({
      dateStr, clockIn: clockTimes[0], clockOut: clockTimes[1],
      actualIn: actualTimes[0], actualOut: actualTimes[1],
      plannedIn: plannedTimes[0], plannedOut: plannedTimes[1],
      status: statusCell, totalTime, breakTime, overtime, nightTime, lateTime, reason
    });
  });

  return days.length > 0 ? { name, formatType, days } : null;
}

/**
 * メイン: ログイン → 月次チェック
 */
function main() {
  const sessionCookie = loginToAkashiWithRetry();
  if (!sessionCookie) {
    sendToGoogleChat(`【勤怠チェック失敗】対象月: ${getTargetYearMonth()} ログイン失敗`, false);
    return;
  }
  checkMonthlyAttendance(sessionCookie);
}

/**
 * 月次勤怠チェック & Chat通知
 */
function checkMonthlyAttendance(sessionCookie) {
  const yearMonth = getTargetYearMonth();
  let cookieMap = cookieStringToMap(sessionCookie);
  const errorsByEmployee = {};
  let totalErrors = 0;
  const { dateStr: todayStr } = getCurrentJSTAndDateString();

  let res = fetchPageHTML(cookieMap, MANAGER_URL);
  if (!res || !res.html || res.html.includes('ログイン')) {
    sendToGoogleChat(`【勤怠チェック失敗】対象月: ${yearMonth} マネージャーページ取得失敗`, false);
    return;
  }
  cookieMap = res.cookieMap;

  const empIds = TARGET_EMPLOYEE_IDS;
  empIds.forEach(empId => {
    const pageUrl = `https://atnd.ak4.jp/ja/manager/attendance/${empId}/${yearMonth}`;
    res = fetchPageHTML(cookieMap, pageUrl);
    if (!res || !res.html || res.html.includes('ログイン')) return;
    cookieMap = res.cookieMap;
    const csrfToken = extractCsrfToken(res.html);
    if (!csrfToken) return;

    const jsText = fetchAttendanceRecords(cookieMap, empId, yearMonth, csrfToken);
    if (!jsText) return;

    const employeeData = parseRecordsFromJs(jsText, res.html, empId);
    if (!employeeData || employeeData.days.length === 0) return;

    const employeeName = employeeData.name;
    if (!errorsByEmployee[employeeName]) errorsByEmployee[employeeName] = [];

    employeeData.days.forEach(day => {
      if (day.dateStr >= todayStr) return;
      const dateStr = day.dateStr;

      let reason = day.reason || '';
      if (reason.startsWith('理由 ')) reason = reason.substring(3).trim();

      let overtime = day.overtime;
      if (employeeData.formatType === 'flex' && isTime(day.totalTime)) {
        overtime = minutesToTime(timeToMinutes(day.totalTime) - 480);
      }

      let errorMessage = '';

      // 1. 打刻忘れ → 一律警告
      if ((!isTime(day.clockIn) && isTime(day.actualIn)) || (!isTime(day.clockOut) && isTime(day.actualOut))) {
        errorMessage = '打刻忘れがあります（理由が記載されていても要確認）';
      }
      // 2. 30分乖離 → 一律警告
      if (isTime(day.clockIn) && isTime(day.actualIn) && Math.abs(timeToMinutes(day.clockIn) - timeToMinutes(day.actualIn)) >= 30) {
        errorMessage = '出勤打刻と実績が30分以上乖離しています（理由が記載されていても要確認）';
      }

      // 3. 遅刻・早退 → 元のキーワードチェック（理由必須）
      if (!errorMessage) {
        if (day.status.includes('遅刻') && !checkReason(reason)) {
          errorMessage = '勤務状況に遅刻があるが理由が記載されていない可能性あり';
        }
        if (day.status.includes('早退') && !checkReason(reason)) {
          errorMessage = '早退しているが理由欄に早退理由が記載されていない可能性あり';
        }
        if (day.status.includes('遅刻') && isTime(day.lateTime) && timeToMinutes(day.lateTime) > 0 && !checkReason(reason)) {
          errorMessage = '遅刻しているが理由欄に遅刻理由が記載されていない可能性あり';
        }
        if (day.status.includes('遅刻') && isTime(day.plannedIn) && isTime(day.actualIn)) {
          const breakMin = timeToMinutes(day.breakTime);
          const lateMin = timeToMinutes(day.lateTime);
          if (lateMin >= 180 && breakMin > 30) errorMessage = '休憩時間が遅刻時間として換算されている可能性あり';
        }
        if (day.status.includes('早退') && isTime(day.plannedOut) && isTime(day.actualOut)) {
          const breakMin = timeToMinutes(day.breakTime);
          const lateMin = timeToMinutes(day.lateTime);
          if (lateMin >= 240 && breakMin > 30) errorMessage = '休憩時間が早退時間として換算されている可能性あり';
        }

        // 4. その他すべてのルール（変更なし）
        if (isTime(day.clockOut) && isTime(day.plannedOut) && timeToMinutes(day.clockOut) < timeToMinutes(day.plannedOut) && !checkReason(reason)) {
          errorMessage = '退勤予定前に打刻されているが理由が記載されていない可能性あり';
        }
        if (reason.includes('在宅勤務') && !day.status.includes('在宅勤務') && !['出社', '出勤', '移動'].some(k => reason.includes(k))) {
          errorMessage = '理由欄に在宅勤務とあるが勤務状況に未記載、出社していなければ勤務状況を更新すること';
        }
        if (day.status.includes('在宅勤務') && ['出社', 'へ出勤', 'に出勤', '移動'].some(k => reason.includes(k))) {
          errorMessage = '勤務状況が在宅勤務とあるが理由欄に出社がある';
        }
        if (isTime(day.totalTime) && isTime(day.breakTime)) {
          let totalMin = timeToMinutes(day.totalTime);
          const breakMin = timeToMinutes(day.breakTime);
          if (day.status.includes('午前半年休')) totalMin -= 180;
          else if (day.status.includes('午後半年休')) totalMin -= 270;
          else if (day.status.match(/年休|記念日休暇/)) totalMin -= 450;
          if (totalMin > 360 && breakMin < 45) errorMessage = '実働6時間超で休憩45分未満';
          if (totalMin > 480 && breakMin < 60) errorMessage = '実働8時間超で休憩1時間未満';
        }
        if (day.status.includes('電車遅延')) {
          if (day.clockIn === day.plannedIn) errorMessage = '電車遅延理由があるが実績出勤時間が予定と同一となっている';
          if (day.status.includes('遅刻')) errorMessage = '電車遅延にも関わらず勤務状況に遅刻が含まれている';
        }
        if (isTime(overtime) && timeToMinutes(overtime) > 0 && (!checkReason(reason) || !reason.includes('指示'))) {
          if (!(employeeData.formatType === 'flex' && timeToMinutes(overtime) <= 60)) {
            errorMessage = '残業があるが理由または指示者が記載されていない';
          }
        }
        if ((day.status.includes('午前半年休') || day.status.includes('午後半年休')) && isTime(overtime) && timeToMinutes(overtime) > 0) {
          errorMessage = '半日休暇にもかかわらず残業が入力されている';
        }
        if ((day.status.includes('午前半年休') || day.status.includes('午後半年休')) && isTime(day.breakTime) && timeToMinutes(day.breakTime) >= 45) {
          errorMessage = '半日休暇にもかかわらず勤務時間外の不要な休憩時間が入力されている可能性あり';
        }
        if (day.status.includes('振替出勤') && (!checkReason(reason) || !reason.includes('指示'))) {
          errorMessage = '振替出勤にもかかわらず、休出理由または指示者が記載されていない';
        }
        if (day.status.includes('振替休日') && !['の振休', 'の振替休日'].some(k => reason.includes(k))) {
          errorMessage = '振替休日にもかかわらず、振替元の日付が記載されていない';
        }
        if (day.status.includes('休日出勤') && (!checkReason(reason) || !reason.includes('指示'))) {
          errorMessage = '休日出勤にもかかわらず、休出理由または指示者が記載されていない';
        }
        if (day.status.includes('代休') && !reason.includes('代休')) {
          errorMessage = '代休にもかかわらず、代休元の日付が記載されていない';
        }
        if (day.status.includes('年休') && !day.status.includes('半年休') && day.status.includes('在宅勤務')) {
          errorMessage = '有休と在宅勤務が同時に記載されている';
        }
      }

      if (errorMessage) {
        errorsByEmployee[employeeName].push({
          dateStr,
          errorMessage,
          status: day.status,
          reason
        });
        totalErrors++;
      }
    });
  });

  // 通知（0件でも送信）
  let message = '';
  if (totalErrors > 0) {
    message = `【勤怠指摘事項_${yearMonth}】 ${totalErrors}件\n`;
    for (const [name, empErrors] of Object.entries(errorsByEmployee)) {
      if (empErrors.length === 0) continue;
      message += `\n対象者=${name} さん\n`;
      empErrors.forEach(err => {
        message += `日付=${err.dateStr}, 指摘=${err.errorMessage}, 勤務状況=${err.status}, 理由欄=${err.reason}\n`;
      });
    }
    sendToGoogleChat(message, true);
  } else {
    message = `【勤怠チェック完了】対象月: ${yearMonth}\n指摘事項なし（0件）です。全て問題ありませんでした。`;
    sendToGoogleChat(message, false);
  }
}

function testJSTDate() {
  const { date: nowJST, dateStr } = getCurrentJSTAndDateString();
  Logger.log(`現在JST: ${nowJST}, dateStr: ${dateStr}`);
}
