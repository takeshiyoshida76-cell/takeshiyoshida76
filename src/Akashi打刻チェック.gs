// 定数: 対象者のリストとURL
const TARGET_NAMES = ['対象者１', '対象者２', '対象者３', '対象者４'];
const LOGIN_URL = 'https://atnd.ak4.jp/ja/login';
const ATTENDANCE_URL = 'https://atnd.ak4.jp/ja/manager/daily_summary';

// ユーティリティ: スクリプトプロパティから認証情報、Webhook URLを取得
function getCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    companyId: scriptProperties.getProperty('COMPANY_ID'),
    loginId: scriptProperties.getProperty('LOGIN_ID'),
    password: scriptProperties.getProperty('PASSWORD'),
    webhookUrl: scriptProperties.getProperty('WEBHOOK_URL')
  };
}

// ユーティリティ: 現在の時刻をJSTで取得
function getCurrentJST() {
  const now = new Date();
  const jstTimeZone = 'Asia/Tokyo';
  const jstDate = Utilities.formatDate(now, jstTimeZone, 'yyyy-MM-dd HH:mm:ss');
  return new Date(jstDate + '+09:00');
}

// ユーティリティ: 時刻文字列をDateオブジェクトに変換
function parseTime(timeStr, baseDate) {
  const [hours, minutes] = timeStr.split(':').map(Number);
  const date = new Date(baseDate);
  date.setHours(hours, minutes, 0, 0);
  return date;
}

// メイン処理: スクリプトのエントリーポイント
function main(e) {
  // イベントタイトルをチェック
  const sessionCookie = loginToAkashi();
  if (sessionCookie) {
    checkAttendance(sessionCookie);
  } else {
    Logger.log('ログインに失敗したため、処理を中断します。');
  }
}

// メイン処理: 勤怠データをチェック
function checkAttendance(sessionCookie) {
  const now = getCurrentJST();
  const htmlContent = fetchAttendanceHTML(sessionCookie);
  const employeeData = parseAttendanceHTML(htmlContent);
  const messages = [];

  employeeData.forEach(employee => {
    if (TARGET_NAMES.includes(employee.name.split(' ')[0])) {
      // 出勤チェック（5分前）
      if (employee.scheduledStart !== '--:--' && employee.actualStart === '--:--') {
        const scheduledStartTime = parseTime(employee.scheduledStart, now);
        const fiveMinutesBefore = new Date(scheduledStartTime.getTime() - 5 * 60 * 1000);
        if (now >= fiveMinutesBefore) {
          messages.push(`${employee.name} さん、予定出勤時刻 ${employee.scheduledStart} の出勤打刻が行われていません。至急、打刻してください。`);
        }
      }
      // 退勤チェック（15分後）
      if (employee.scheduledEnd !== '--:--' && employee.actualEnd === '--:--') {
        const scheduledEndTime = parseTime(employee.scheduledEnd, now);
        const fifteenMinutesAfter = new Date(scheduledEndTime.getTime() + 15 * 60 * 1000);
        if (now >= fifteenMinutesAfter) {
          messages.push(`${employee.name} さん、予定退勤時刻 ${employee.scheduledEnd} の退勤打刻が行われていません。至急、打刻してください。`);
        }
      }
    }
  });

  if (messages.length > 0) {
    const messageText = messages.join('\n');
    sendToGoogleChat(messageText);
    Logger.log('送信したメッセージ:\n' + messageText);
  } else {
    Logger.log('送信するメッセージはありません。');
  }
}

// 外部通信: Akashiにログインし、セッションCookieを取得
function loginToAkashi() {
  const credentials = getCredentials();
  const csrfToken = getCsrfToken();

  if (!csrfToken) {
    Logger.log('ログイン中止: CSRFトークンが取得できませんでした。');
    return null;
  }

  const payload = {
    'authenticity_token': csrfToken,
    'form[company_id]': credentials.companyId,
    'form[login_id]': credentials.loginId,
    'form[password]': credentials.password,
    'form[next]': '/ja/manager',
    'form[fill_company_id_and_login_id]': '1',
    'commit': 'ログイン'
  };

  const options = {
    method: 'POST',
    payload: payload,
    followRedirects: false,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(LOGIN_URL, options);
    const statusCode = response.getResponseCode();
    const headers = response.getAllHeaders();

    if (statusCode === 302) {
      const cookies = headers['Set-Cookie'] || headers['set-cookie'];
      if (cookies) {
        let sessionCookie = '';
        cookies.forEach(cookie => {
          sessionCookie += cookie.split(';')[0] + '; ';
        });
        Logger.log('ログイン成功、Cookie: ' + sessionCookie);
        return sessionCookie.trim();
      } else {
        Logger.log('エラー: Cookieが取得できませんでした。');
        return null;
      }
    } else {
      Logger.log('ログイン失敗、ステータスコード: ' + statusCode);
      Logger.log('レスポンス: ' + response.getContentText());
      return null;
    }
  } catch (e) {
    Logger.log('エラー: ログイン処理に失敗: ' + e.message);
    return null;
  }
}

// 外部通信: ログイン画面からCSRFトークンを取得
function getCsrfToken() {
  const options = {
    method: 'GET',
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(LOGIN_URL, options).getContentText();
    const match = response.match(/<input[^>]*name="authenticity_token"[^>]*value="([^"]*)"/);
    if (match && match[1]) {
      Logger.log('CSRFトークンを取得: ' + match[1]);
      return match[1];
    } else {
      Logger.log('エラー: CSRFトークンが見つかりませんでした。');
      return null;
    }
  } catch (e) {
    Logger.log('エラー: CSRFトークンの取得に失敗: ' + e.message);
    return null;
  }
}

// 外部通信: 勤務表のHTMLを取得
function fetchAttendanceHTML(sessionCookie) {
  try {
    const options = {
      headers: {
        'Cookie': sessionCookie
      },
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(ATTENDANCE_URL, options);
    if (response.getResponseCode() === 200) {
      return response.getContentText('UTF-8');
    } else {
      Logger.log('HTMLの取得に失敗しました: ' + response.getResponseCode());
      return '';
    }
  } catch (e) {
    Logger.log('エラー: HTML取得に失敗: ' + e.message);
    return '';
  }
}

// 外部通信: Google Chatにメッセージを送信
function sendToGoogleChat(message) {
  const credentials = getCredentials();
  const webhookUrl = credentials.webhookUrl;

  if (!webhookUrl) {
    Logger.log('エラー: WEBHOOK_URLがスクリプトプロパティに設定されていません。');
    return;
  }

  const payload = {
    text: message
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('Google Chatへの送信エラー: ステータスコード ' + response.getResponseCode());
    }
  } catch (e) {
    Logger.log('Google Chatへの送信エラー: ' + e.message);
  }
}

// データ解析: HTMLを解析して従業員データを抽出
function parseAttendanceHTML(html) {
  const employees = [];
  const rowRegex = /<tr id="daily_summary_\d+"[\s\S]*?<\/tr>/g;
  const rows = html.match(rowRegex) || [];

  rows.forEach(row => {
    const nameMatch = row.match(/<ruby[^>]*>([^<]+)<rt>[^<]+<\/rt><\/ruby><ruby[^>]*>([^<]+)<rt>[^<]+<\/rt><\/ruby>/);
    const name = nameMatch ? `${nameMatch[1]} ${nameMatch[2]}` : '';

    if (name && TARGET_NAMES.includes(name.split(' ')[0])) {
      const scheduledMatch = row.match(/<td class="c-main-table-body__cell">[\s\S]*?<div>\s*<span class="[^"]*">([^<]+)<\/span><br\/>\s*<span class="[^"]*">([^<]+)<\/span><br\/>\s*<\/div>/g);
      const actualMatch = row.match(/<td class="c-main-table-body__cell">[\s\S]*?<div>\s*<span class="">([^<]+)<\/span><br\/>\s*<span class="">([^<]+)<\/span><br\/>\s*<\/div>/g);

      let scheduledStart = '--:--';
      let scheduledEnd = '--:--';
      let actualStart = '--:--';
      let actualEnd = '--:--';

      if (scheduledMatch && scheduledMatch[2]) {
        const times = scheduledMatch[2].match(/<span class="[^"]*">([^<]+)<\/span><br\/>\s*<span class="[^"]*">([^<]+)<\/span>/);
        if (times) {
          scheduledStart = times[1];
          scheduledEnd = times[2];
        }
      }

      if (actualMatch && actualMatch[0]) {
        const times = actualMatch[0].match(/<span class="">([^<]+)<\/span><br\/>\s*<span class="">([^<]+)<\/span>/);
        if (times) {
          actualStart = times[1];
          actualEnd = times[2];
        }
      }

      employees.push({
        name: name,
        scheduledStart: scheduledStart,
        scheduledEnd: scheduledEnd,
        actualStart: actualStart,
        actualEnd: actualEnd
      });
    }
  });

  return employees;
}

// テスト: ログイン確認用
function testLogin() {
  const sessionCookie = loginToAkashi();
  if (sessionCookie) {
    Logger.log('セッション Cookie: ' + sessionCookie);
  } else {
    Logger.log('ログインに失敗しました。');
  }
}
