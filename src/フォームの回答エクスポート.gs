/**
 * @fileoverview Googleフォームの新しい回答を、フォームと同じディレクトリにCSVとしてエクスポートします。
 * スクリプトはフォームIDを直接指定して動作する単独のスクリプトとして設計されています。
 */

const FORM_ID = PropertiesService.getScriptProperties().getProperty('BP_FORM_ID');

// スクリプトがGoogleドライブに保存するフォルダの名前
const FOLDER_NAME = 'フォーム回答バックアップ';

/**
 * フォームの回答をチェックし、新しい回答があった場合のみCSVとしてエクスポートします。
 */
function exportNewFormResponsesToCsv() {
  try {
    // 指定されたフォームIDを使用してフォームを開く
    const form = FormApp.openById(FORM_ID);

    if (!form) {
      Logger.log('エラー: 指定されたフォームIDが見つかりません。');
      return;
    }

    // スクリプトのユーザープロパティから最終実行日時を取得
    const scriptProperties = PropertiesService.getScriptProperties();
    const lastRunTimeStr = scriptProperties.getProperty('lastRunTime');
    const lastRunTime = lastRunTimeStr ? new Date(lastRunTimeStr) : new Date(0);

    // フォームの全回答を取得
    const formResponses = form.getResponses();
    if (formResponses.length === 0) {
      Logger.log('回答がありません。エクスポートをスキップします。');
      return;
    }

    // 前回の実行日時より新しい回答だけをフィルタリング
    const newResponses = formResponses.filter(response => {
      const timestamp = response.getTimestamp();
      return timestamp > lastRunTime;
    });

    if (newResponses.length === 0) {
      Logger.log('新しい回答はありません。エクスポートをスキップします。');
      return;
    }

    // ヘッダー行を作成
    const items = form.getItems();
    const headers = items.map(item => `"${item.getTitle().replace(/"/g, '""')}"`).join(',');

    // 新しい回答データをCSV形式の文字列に変換
    let csvContent = headers + '\n';
    newResponses.forEach(formResponse => {
      const itemResponses = formResponse.getItemResponses();
      
      const row = itemResponses.map(itemResponse => {
        const answer = itemResponse.getResponse();
        let responseValue = Array.isArray(answer) ? answer.join(', ') : answer;
        responseValue = `"${responseValue.toString().replace(/"/g, '""')}"`;
        return responseValue;
      }).join(',');
      
      csvContent += row + '\n';
    });

    // --- ここから変更点 ---
    // フォームの親フォルダを取得
    const formFile = DriveApp.getFileById(form.getId());
    const parentFolders = formFile.getParents();
    
    // 親フォルダがない場合はマイドライブを使用
    let parentFolder;
    if (parentFolders.hasNext()) {
      parentFolder = parentFolders.next();
    } else {
      parentFolder = DriveApp.getRootFolder();
    }

    // フォームの親フォルダ内に、指定されたフォルダが存在するか確認し、なければ作成
    let targetFolder;
    const folders = parentFolder.getFoldersByName(FOLDER_NAME);
    if (folders.hasNext()) {
      targetFolder = folders.next();
    } else {
      targetFolder = parentFolder.createFolder(FOLDER_NAME);
    }
    // --- 変更点ここまで ---

    // ファイル名にタイムスタンプを追加
    const formTitle = form.getTitle().replace(/ /g, '_');
    const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const fileName = `${formTitle}_${formattedDate}_new_responses.csv`;

    // CSV文字列をファイルとして対象フォルダに保存
    targetFolder.createFile(fileName, csvContent, MimeType.PLAIN_TEXT);

    Logger.log(`${newResponses.length}件の新しい回答をCSVとして正常にエクスポートしました。`);

    // 最終実行日時をプロパティに保存
    scriptProperties.setProperty('lastRunTime', new Date().toISOString());

  } catch (error) {
    Logger.log('CSVエクスポート中にエラーが発生しました: ' + error.toString());
  }
}
