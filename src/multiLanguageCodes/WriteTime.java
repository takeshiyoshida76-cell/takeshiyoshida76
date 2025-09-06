import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * A Java class to write a timestamped log message to a file.
 */
public class WriteTime {

    // Use a constant for the filename for easy maintenance
    private static final String FILENAME = "MYFILE.txt";

    /**
     * The main method is the entry point of the application.
     * It writes a program message and the current system time to a log file.
     * @param args Command-line arguments (not used).
     */
    public static void main(String[] args) {

        // Define the date and time format
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
        
        // Get the current date and time
        LocalDateTime now = LocalDateTime.now();
        String nowtime = now.format(formatter);

        // Use a try-with-resources block to ensure the writer is automatically closed
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(FILENAME, true))) {
            
            // Write a program message, followed by a new line
            writer.write("This program is written in Java.");
            writer.newLine();

            // Write the formatted system time, followed by a new line
            writer.write("Current Time = " + nowtime);
            writer.newLine();

            // Print a success message to standard output
            System.out.println("Successfully appended log to " + FILENAME);

        } catch (IOException e) {
            // Print a user-friendly error message to the standard error stream
            System.err.println("ERROR: An error occurred while writing to the file.");
            System.err.println("Details: " + e.getMessage());
        }
    }
}
