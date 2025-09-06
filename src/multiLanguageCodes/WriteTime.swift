import Foundation

func appendToFile(message: String, filename: String) {
    let fileURL = URL(fileURLWithPath: filename)

    do {
        // ファイルが存在するか確認
        if FileManager.default.fileExists(atPath: fileURL.path) {
            // ファイルが存在すれば、データを追記
            if let fileHandle = try? FileHandle(forUpdating: fileURL) {
                fileHandle.seekToEndOfFile()
                if let data = message.data(using: .utf8) {
                    fileHandle.write(data)
                }
                fileHandle.closeFile()
            }
        } else {
            // ファイルが存在しない場合は、新規作成して書き込み
            try message.write(to: fileURL, atomically: true, encoding: .utf8)
        }
    } catch {
        print("ファイルへの書き込み中にエラーが発生しました: \(error.localizedDescription)")
    }
}

// ----------------------------------------------------
// メイン処理
// ----------------------------------------------------

let filename = "MYFILE.txt"

// 現在時刻の取得とフォーマット
let now = Date()
let formatter = DateFormatter()
formatter.dateFormat = "yyyy/MM/dd HH:mm:ss"
let nowtime = formatter.string(from: now)

// ログメッセージの作成
let logMessage = """
This program is written in Swift.
Current Time = \(nowtime)

"""

// ログをファイルに追記
appendToFile(message: logMessage, filename: filename)

