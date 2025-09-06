import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class WriteTime {
    public static void main(String[] args) {
        String filename = "MYFILE.TXT";

        // DateTime Format
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
        LocalDateTime now = LocalDateTime.now();
        String nowtime = now.format(formatter);

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filename, true))) {
            // Write Message
            writer.write("This program is written in TAL.");
            writer.newLine();

            // Write SystemDateTime
            writer.write("Current Time = " + nowtime);
            writer.newLine();

         } catch (IOException e) {
            e.printStackTrace();
         }
     }
}
