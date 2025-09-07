# -*- coding: utf-8 -*-
from datetime import datetime

OUTPUT_FILE = "MYFILE.txt"

def main():
    # Get the current date and time in YYYY/MM/DD HH:MM:SS format
    now = datetime.now()
    now_str = now.strftime("%Y/%m/%d %H:%M:%S")

    # Write the content to the file (overwrite)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("This program was written in Python.\n")
        f.write(f"Current time is: {now_str}\n")

    print(f"Successfully wrote current time to '{OUTPUT_FILE}'.")

if __name__ == "__main__":
    main()
