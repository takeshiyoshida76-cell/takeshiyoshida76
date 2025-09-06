import java.io.File
import java.io.IOException
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

/**
 * A Groovy script to append a log message with the current date and time
 * to a specified file.
 */
class WriteTime {

    private static final String FILENAME = 'MYFILE.txt'

    static void main(String[] args) {
        def file = new File(FILENAME)

        // Get current system datetime and format it
        def nowtime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"))
        
        def logMessage = "This program is written in Groovy.\nCurrent Time = ${nowtime}\n"

        // Use a try-catch block for robust error handling
        try {
            // Append the log message to the file. 'withWriterAppend' automatically
            // handles file opening, closing, and resource management.
            file.withWriterAppend { writer ->
                writer.write(logMessage)
            }
            println("Log message successfully appended to: ${FILENAME}")
        } catch (IOException e) {
            // Print an error message if an IOException occurs during file operations
            System.err.println("ERROR: An error occurred while writing to the file.")
            System.err.println("Exception details: ${e.message}")
        }
    }
}
