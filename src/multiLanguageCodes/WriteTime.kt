import java.nio.file.Files
import java.nio.file.Paths
import java.nio.file.StandardOpenOption
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

fun main() {
    // ファイル名
    val filename = "MYFILE.TXT"

    // 現在時刻の取得とフォーマット
    val now = LocalDateTime.now()
    val formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss")
    val nowtime = now.format(formatter)

    // ログメッセージの作成
    val message = "This program is written in TAL.\nCurrent Time = $nowtime\n"

    try {
        // ファイルに追記。ファイルが存在しない場合は新規作成。
        Files.write(
            Paths.get(filename),
            message.toByteArray(),
            StandardOpenOption.CREATE,
            StandardOpenOption.APPEND
        )
    } catch (e: Exception) {
        println("ファイルへの書き込み中にエラーが発生しました: ${e.message}")
    }
}
