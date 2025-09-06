import java.io.File
import java.io.IOException
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

def filename = 'MYFILE.txt'
def file = new File(filename)

// Get System Datetime
def nowtime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"))

def logMessage = "This program is written in TAL.\nCurrent Time = ${nowtime}\n"

try {
    file.withWriterAppend { writer ->
        writer.write(logMessage)
    }
} catch (IOException e) {
    println("ファイルへの書き込み中にエラーが発生しました: ${e.message}")
}
