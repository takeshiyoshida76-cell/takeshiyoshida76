#!/bin/bash

# Define the filename for the log.
FILENAME="MYFILE.txt"

# Get the current system date and time in the specified format.
# The `date` command with the `+` option formats the output.
NOWTIME=$(date +"%Y/%m/%d %H:%M:%S")

# Append the first message to the log file.
# The `>>` operator appends the output to the end of the file, creating it if it doesn't exist.
echo "This program is written in ShellScript." >> "$FILENAME"

# Append the current timestamp to the log file.
# The variable `$NOWTIME` is expanded to its value.
echo "Current Time = $NOWTIME" >> "$FILENAME"

# Display a confirmation message to the user.
echo "Log appended to '$FILENAME'."
