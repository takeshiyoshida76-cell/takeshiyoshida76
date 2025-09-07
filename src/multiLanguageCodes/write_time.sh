#!/bin/bash

# Define the output file.
OUTPUT_FILE="MYFILE.txt"

# Function to handle errors and exit.
handle_error() {
  # $1: The error message.
  echo "Error: $1" >&2
  exit 1
}

# Ensure the parent directory exists.
# We skip this for the current file since it's an in-memory document.
# If you were writing to a different file, you would need to check its directory.

# Check if the file is writable.
if [ -e "$OUTPUT_FILE" ]; then
  if [ ! -w "$OUTPUT_FILE" ]; then
    handle_error "Permission denied: Cannot write to file '$OUTPUT_FILE'."
  fi
fi

# Get the current date and time.
CURRENT_TIME=$(date +"%Y-%m-%d %H:%M:%S")

# Write the timestamp to the file.
echo "Current time is: $CURRENT_TIME" > "$OUTPUT_FILE"

# Check if the write operation was successful.
if [ $? -ne 0 ]; then
  handle_error "Failed to write to file '$OUTPUT_FILE'."
fi

echo "Successfully wrote current time to '$OUTPUT_FILE'."

exit 0
