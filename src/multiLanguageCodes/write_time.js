const fs = require('fs');

// ----------------------------------------------------
// メイン処理
// ----------------------------------------------------

/**
 * 現在時刻を指定されたフォーマットで文字列化する関数
 * @returns {string} フォーマットされた時刻文字列 (例: '2023/10/27 15:30:00')
 */
const getFormattedCurrentTime = () => {
    const now = new Date();
    const pad = (num) => String(num).padStart(2, '0');

    const year = now.getFullYear();
    const month = pad(now.getMonth() + 1); // getMonth()は0から始まるため+1
    const day = pad(now.getDate());
    const hours = pad(now.getHours());
    const minutes = pad(now.getMinutes());
    const seconds = pad(now.getSeconds());

    return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
};

const filename = 'MYFILE.TXT';
const logMessage = `This program is written in JavaScript.\nCurrent Time = ${getFormattedCurrentTime()}\n`;

try {
    // ファイルにメッセージを追記
    // appendFileSyncは同期的にファイルを追記するので、処理完了まで待機
    fs.appendFileSync(filename, logMessage);
    console.log(`ログが'${filename}'に追記されました。`);
} catch (err) {
    // ファイル操作でエラーが発生した場合
    console.error('ファイルへの書き込み中にエラーが発生しました:', err);
}
