# -*- coding: utf-8 -*-
import logging
from datetime import datetime

# Set up basic logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

def write_log(filename: str, message: str) -> None:
    """
    Appends a given message with a timestamp to a file.

    Args:
        filename (str): The name of the file to write to.
        message (str): The message to append to the file.
    """
    try:
        # Use 'a' for append mode to add content without overwriting.
        # 'utf-8' encoding is specified for wide character support.
        with open(filename, "a", encoding="utf-8") as f:
            f.write(message)
        logging.info("Successfully appended log to '%s'.", filename)
    except IOError as e:
        # Log an error if file operation fails.
        logging.error("An error occurred while writing to the file: %s", e)

if __name__ == "__main__":
    # Define the log filename.
    filename = "MYFILE.txt"

    # Get the current date and time.
    now = datetime.now()
    nowtime_str = now.strftime("%Y/%m/%d %H:%M:%S")
    
    # Create the log message.
    log_message = f"This program is written in Python.\nCurrent Time = {nowtime_str}\n"

    # Write the log message to the file.
    write_log(filename, log_message)
