// @ts-nocheck
/**
 * Akashi 月間残業15時間超アラート自動化スクリプト (Google Apps Script)
 *
 * 【概要】
 * - AKASHIの管理者ページから「月間残業時間」「月間所定休日労働時間」「月間法定休日労働時間」を取得
 * - 3つの合計が15時間以上の社員を検知 → Google Chatに@all通知
 *
 * 【使用方法】
 * 1. スクリプトプロパティに以下を設定（既存と同じ）
 * COMPANY_ID, LOGIN_ID, PASSWORD, CHAT_WEBHOOK_URL
 * 2. TARGET_EMPLOYEE_IDS を監視したい従業員IDに変更
 * 3. トリガーで main() を毎日or毎月25日などに登録 → 完了！
 */

// ==============================================================================
// 1. 定数設定
// ==============================================================================

const TARGET_EMPLOYEE_IDS = [336546, 336415, 336530, 336540]; // 監視対象従業員ID
const THRESHOLD_HOURS = 15;
const THRESHOLD_MINUTES = THRESHOLD_HOURS * 60;

const LOGIN_URL = 'https://atnd.ak4.jp/ja/login';
const MANAGER_URL = 'https://atnd.ak4.jp/ja/manager';
const ROOT_JA_URL = 'https://atnd.ak4.jp/ja';
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36';

// ==============================================================================
// 2. ユーティリティ関数（既存スクリプトと100%互換）
// ==============================================================================

/**
 * スクリプトプロパティから認証情報を取得します。
 * 必要なプロパティが設定されていない場合はエラーをスローします。
 * @returns {{companyId: string, loginId: string, password: string}} 認証情報オブジェクト
 * @throws {Error} 認証情報が設定されていない場合
 */
function getCredentials() {
  const companyId = PropertiesService.getScriptProperties().getProperty('COMPANY_ID');
  const loginId = PropertiesService.getScriptProperties().getProperty('LOGIN_ID');
  const password = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (!companyId || !loginId || !password) throw new Error('認証情報が設定されていません。');
  return { companyId, loginId, password };
}

/**
 * 指定された範囲でランダムな遅延を発生させます。
 * @param {number} minMs 最小遅延時間 (ミリ秒)
 * @param {number} maxMs 最大遅延時間 (ミリ秒)
 */
function addRandomDelay(minMs, maxMs) {
  const delay = Math.floor(Math.random() * (maxMs - minMs + 1)) + minMs;
  Utilities.sleep(delay);
}

/**
 * HTTPレスポンスの 'Set-Cookie' ヘッダーからCookieマップを更新します。
 * @param {Object.<string, string>} cookieMap 現在のCookieマップ
 * @param {string|string[]|null} setCookieHeaders 'Set-Cookie' ヘッダーの値
 * @returns {Object.<string, string>} 更新されたCookieマップ
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
 * CookieマップをHTTPヘッダー形式の文字列にフォーマットします。
 * @param {Object.<string, string>} cookieMap Cookieマップ
 * @returns {string} フォーマットされたCookieヘッダー文字列
 */
function formatCookieHeader(cookieMap) {
  return Object.entries(cookieMap).map(([k, v]) => `${k}=${v}`).join('; ');
}

/**
 * Cookie文字列をCookieマップに変換します。
 * @param {string} cookieStr Cookie文字列
 * @returns {Object.<string, string>} Cookieマップ
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
 * Google Chatにメッセージを送信します。
 * @param {string} message 送信するメッセージ本文
 * @param {boolean} [includeMention=true] @all を含めるかどうか
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
 * 監視対象となる年と月をYYYYMM形式で取得します。
 * 22日以降は翌月を対象とします。
 * @returns {string} YYYYMM形式の年月文字列
 */
function getTargetYearMonth() {
  const now = new Date();
  const day = now.getDate();
  let year = now.getFullYear();
  let month = now.getMonth() + 1;
  if (day >= 22) {
    month += 1;
    if (month > 12) { month = 1; year += 1; }
  }
  return `${year}${month.toString().padStart(2, '0')}`;
}

/**
 * HH:MM形式の時間を分単位に変換します。
 * @param {string} str HH:MM形式の時刻文字列
 * @returns {number} 分単位に変換された時間、または無効な場合は0
 */
function timeToMinutes(str) {
  if (!str || !/^\d+:\d{2}$/.test(str)) return 0;
  const [h, m] = str.split(':').map(Number);
  return h * 60 + m;
}

/**
 * 分単位の時間をHH:MM形式に変換します。
 * @param {number} mins 分単位の時間
 * @returns {string} HH:MM形式の時刻文字列
 */
function minutesToTime(mins) {
  if (mins < 0) mins = 0;
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h}:${m.toString().padStart(2, '0')}`;
}

/**
 * HTMLからCSRFトークンを抽出します。
 * @param {string} html ページのHTMLコンテンツ
 * @returns {string|null} CSRFトークン、または見つからない場合はnull
 */
function extractCsrfToken(html) {
  const match = html.match(/<meta name="csrf-token" content="([^"]+)" \/>/);
  return match ? match[1] : null;
}

/**
 * 指定されたURLからHTMLコンテンツをフェッチします。
 * ログインセッションの維持とリトライロジックを含みます。
 * @param {Object.<string, string>} initialCookieMap 初期Cookieマップ
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

      if (html.includes('<title>AKASHI - ログイン</title>')) {
        const newCookie = loginToAkashiWithRetry();
        if (newCookie) cookieMap = cookieStringToMap(newCookie);
        retryCount++;
        continue;
      }
      if (response.getResponseCode() === 200) return { html, cookieMap };
    } catch (e) {
      retryCount++;
    }
  }
  return { html: '', cookieMap };
}

/**
 * AKASHIログイン処理を実行します。
 * @return {string|null} 最終Cookie文字列、またはログイン失敗時はnull
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
 * AKASHIログイン処理を複数回リトライするラッパー関数。
 * @param {number} [maxRetries=3] 最大リトライ回数
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
 * @param {Object.<string, string>} cookieMap 現在のCookieマップ
 * @param {number} empId 従業員ID
 * @param {string} yearMonth YYYYMM形式の年月文字列
 * @returns {string} サマリーデータのレスポンステキスト
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
 * @param {string} text AKASHIから取得したサマリーデータテキスト（HTML断片）
 * @returns {number} 合計時間（分単位）
 */
function parseTotalOvertimeMinutes(text) {
  let total = 0;

  // 修正後の正規表現リスト
  // ポイント:
  // 1. [\s\S]*? を .*? に変更し、非貪欲マッチを維持。
  // 2. 正規表現がマッチしない原因であるタグ間のエスケープ文字や改行を許容するため、
  //    時間抽出部分の直前にある HTML/エスケープ文字を .*? で広く許容する。
  const regs = [
    // 1. 月間残業時間: 「月間残業時間」から「時間:分」までをワイルドカードで繋ぐ
    /月間残業時間.*?p-summary-roster-table-body__value[^>]*>(?:\\n|[\s\S])*?(\d+):(\d{2})/m,

    // 2. 月間所定休日労働時間
    /月間所定休日労働時間.*?p-summary-roster-table-body__value[^>]*>(?:\\n|[\s\S])*?(\d+):(\d{2})/m,

    // 3. 月間法定休日労働時間
    /月間法定休日労働時間.*?p-summary-roster-table-body__value[^>]*>(?:\\n|[\s\S])*?(\d+):(\d{2})/m
  ];

  regs.forEach((re, i) => {
    const m = text.match(re);

    // 入力データには「月間残業時間」に '0:37' が含まれているため、m は非nullになるはずです。
    // text.match(/月間残業時間/g) などで確認してみると良いでしょう。

    if (m) {
      const h = parseInt(m[1], 10);
      const mm = parseInt(m[2], 10);
      total += h * 60 + mm;
      const names = ['残業', '所定休日', '法定休日']; // Loggerで使用するが、現在のコードではLogger.logはコメントアウトされているため実質未使用
      // Logger.log(`検出 → 月間${names[i]}労働時間: ${h}:${mm.toString().padStart(2,'0')}`); // 元のコードでコメントアウトされているためそのまま
    }
  });
  return total;
}

// ==============================================================================
// メイン関数（これだけトリガーに登録）
// ==============================================================================

/**
 * メイン処理を実行します。
 * 監視対象従業員の月間時間外労働時間を取得し、閾値を超過した場合はGoogle Chatに通知します。
 * この関数をGoogle Apps Scriptのトリガーに登録して使用します。
 */
function main() {
  const yearMonth = getTargetYearMonth();
  const disp = `${yearMonth.slice(0, 4)}年${parseInt(yearMonth.slice(4))}月`;
  Logger.log(`【時間外労働15時間以上チェック開始】${disp}`);

  const cookie = loginToAkashiWithRetry();
  if (!cookie) {
    sendToGoogleChat(`【失敗】${disp} ログインできませんでした`, false);
    return;
  }

  let cookieMap = cookieStringToMap(cookie);
  const alerts = [];
  const results = [];

  for (const empId of TARGET_EMPLOYEE_IDS) {
    addRandomDelay(3000, 7000);

    const page = fetchPageHTML(cookieMap, `https://atnd.ak4.jp/ja/manager/attendance/${empId}/${yearMonth}`);
    if (!page.html.includes('p-roster-header__name')) continue; // ページが正常に取得できなかった、または社員名が含まれない場合はスキップ
    cookieMap = page.cookieMap;

    const nameRaw = page.html.match(/<h2 class="p-roster-header__name">([\s\S]*?)<\/h2>/)?.[1] || '';
    const name = nameRaw.replace(/<rt>[\s\S]*?<\/rt>/gi, '').replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();

    const summaryText = fetchSummaryData(cookieMap, empId, yearMonth);
    if (!summaryText) {
      results.push({ name, time: '取得失敗' });
      continue;
    }

    const mins = parseTotalOvertimeMinutes(summaryText);
    const timeStr = minutesToTime(mins);
    Logger.log(`${name} → 月間時間外労働時間: ${timeStr}`);
    results.push({ name, time: timeStr });

    if (mins >= THRESHOLD_MINUTES) alerts.push({ name, time: timeStr });
  }

  if (alerts.length > 0) {
    let msg = `【月間時間外労働15時間以上アラート】${disp}\n\n`;
    alerts.forEach(a => msg += `${a.name} さん → ${a.time}\n`);
    msg += `\n※残業＋所定休日労働＋法定休日労働の合計です`;
    sendToGoogleChat(msg, true);
    Logger.log('Chat通知送信完了');
  } else {
    Logger.log('該当者なし → 通知なし');
  }

  Logger.log('【チェック完了】');
}
