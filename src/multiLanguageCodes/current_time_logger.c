#include <stdio.h>
#include <time.h>

int main() {
    FILE *outfile;
    time_t t;
    struct tm *tm_info;
    char nowtime[20]; // YYYY/MM/DD HH:MM:SS

    // Open Outfile
    outfile = fopen("MYFILE.TXT", "a");
    if (outfile == NULL) {
        perror("Failed to open file");
        return 1;
    }

    // Write message
    fprintf(outfile, "This program is written in C.\n");

    // Get system datetime
    t = time(NULL);
    tm_info = localtime(&t);

    // Edit message
    strftime(nowtime, sizeof(nowtime), "%Y/%m/%d %H:%M:%S", tm_info);

    // Write message
    fprintf(outfile, "Current Time = %s\n", nowtime);

    // Close Outfile
    fclose(outfile);

     return 0;
}
