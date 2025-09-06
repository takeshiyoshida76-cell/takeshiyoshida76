#!/bin/bash

FILENAME="MYFILE.txt"

# Get System Datetime
NOWTIME=$(date +"%Y/%m/%d %H:%M:%S")

# Write Message
echo "This program is written in ShellScript." >> "$FILENAME"
echo "Current Time = $NOWTIME" >> "$FILENAME"
