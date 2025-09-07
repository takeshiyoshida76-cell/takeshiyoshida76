#!/bin/bash

# Typical ShellScript for writing current time to a file

OUTPUT_FILE="MYFILE.txt"

# Get the current date and time (YYYY-MM-DD HH:MM:SS)
CURRENT_TIME="$(date '+%Y-%m-%d %H:%M:%S')"

# Write the content to the file (overwrite)
echo "This program was written in ShellScript." > "$OUTPUT_FILE"
echo "Current time is: $CURRENT_TIME" >> "$OUTPUT_FILE"

# Check if the file was successfully written
if [ -f "$OUTPUT_FILE" ]; then
  echo "Successfully wrote current time to '$OUTPUT_FILE'."
else
  echo "Error: Failed to write to file '$OUTPUT_FILE'." >&2
  exit 1
fi

exit 0
