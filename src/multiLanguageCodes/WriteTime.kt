import java.nio.file.Files
import java.nio.file.Paths
import java.nio.file.StandardOpenOption
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

/**
 * A Kotlin application to write a timestamped log message to a file.
 */
class WriteTime {

    companion object {
        // Use a constant for the filename for easy maintenance
        private const val FILENAME = "MYFILE.txt"
    }

    fun run() {
        // Get the current system datetime and format it
        val now = LocalDateTime.now()
        val formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss")
        val nowtime = now.format(formatter)

        // Create the log message using a template string
        val message = "This program is written in Kotlin.\nCurrent Time = $nowtime\n"

        // Use a try-catch block for robust error handling
        try {
            // Write to the file. This automatically creates the file if it doesn't exist
            // and appends to it if it does.
            Files.write(
                Paths.get(FILENAME),
                message.toByteArray(),
                StandardOpenOption.CREATE,
                StandardOpenOption.APPEND
            )
            println("Successfully appended log to $FILENAME")
        } catch (e: Exception) {
            // Print a user-friendly error message to the standard error stream
            System.err.println("ERROR: An error occurred while writing to the file.")
            System.err.println("Details: ${e.message}")
        }
    }
}

/**
 * The entry point of the application.
 */
fun main() {
    WriteTime().run()
}
