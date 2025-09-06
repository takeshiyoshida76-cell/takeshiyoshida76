#include <stdio.h>
#include <time.h>
#include <string.h>
#include <stdbool.h>

#define FILENAME "MYFILE.txt"
#define TIMESTAMP_BUFFER_SIZE 20

/**
 * @brief Appends a timestamped log message to a file.
 *
 * @param filename The path to the file.
 * @return Returns true on success, false on failure.
 */
bool writeTime(const char* filename) {
    FILE* outfile = NULL;
    time_t t;
    struct tm* tm_info;
    char nowtime[TIMESTAMP_BUFFER_SIZE];

    // Open file in append mode.
    outfile = fopen(filename, "a");
    if (outfile == NULL) {
        perror("Failed to open file");
        return false;
    }

    // Get the current system date and time.
    t = time(NULL);
    tm_info = localtime(&t);
    if (tm_info == NULL) {
        fprintf(stderr, "Failed to get local time.\n");
        fclose(outfile);
        return false;
    }

    // Format the timestamp string.
    strftime(nowtime, TIMESTAMP_BUFFER_SIZE, "%Y/%m/%d %H:%M:%S", tm_info);

    // Write the messages to the file.
    fprintf(outfile, "This program is written in C.\n");
    fprintf(outfile, "Current Time = %s\n\n", nowtime);

    // Close the file.
    fclose(outfile);
    return true;
}

int main() {
    // Attempt to write the log and check the return value.
    if (writeTime(FILENAME)) {
        printf("Successfully appended log to %s.\n", FILENAME);
    } else {
        // Error message is already printed inside writeLog.
        printf("Program terminated with an error.\n");
        return 1;
    }

    return 0;
}
