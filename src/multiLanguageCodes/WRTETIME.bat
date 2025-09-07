@echo off

rem This is a Windows Batch script version.

rem Define the output file.
set "OUTPUT_FILE=MTFILE.txt"

rem Set a variable for the current date and time.
rem The format is YYYY-MM-DD HH:MM:SS.
for /f "tokens=1-4" %%a in ('date /T') do (
    set "DATE_TEMP=%%c-%%b-%%d"
)
for /f "tokens=1-2" %%a in ('time /T') do (
    set "TIME_TEMP=%%a %%b"
)
set "CURRENT_TIME=%DATE_TEMP% %TIME_TEMP%"

rem Write the content to the file.
echo This program is written in Windows CommandPropmpt's Batch. > "%OUTPUT_FILE%"
echo Current time is: %CURRENT_TIME% >> "%OUTPUT_FILE%"

rem Check if the file was successfully written.
if exist "%OUTPUT_FILE%" (
    echo Successfully wrote current time to "%OUTPUT_FILE%".
) else (
    echo Error: Failed to write to file "%OUTPUT_FILE%".
exit /b 1
)

exit /b 0
