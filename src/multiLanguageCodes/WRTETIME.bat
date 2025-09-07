@echo off
rem This is a typical Windows Batch script version.

rem Define the output file.
set "OUTPUT_FILE=MYFILE.txt"

rem Get current date and time (as-is, for standard batch usage)
set "CURRENT_TIME=%date% %time%"

rem Write the content to the file.
echo This program was written in Windows Batch. > "%OUTPUT_FILE%"
echo Current time is: %CURRENT_TIME% >> "%OUTPUT_FILE%"

rem Check if the file was successfully written.
if exist "%OUTPUT_FILE%" (
    echo Successfully wrote current time to "%OUTPUT_FILE%".
) else (
    echo Error: Failed to write to file "%OUTPUT_FILE%".
    exit /b 1
)

exit /b 0
