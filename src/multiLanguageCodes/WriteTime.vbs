' This script appends a timestamped log message to a file.
' The file is automatically created if it does not exist.

' --- Constants ---
Const FOR_APPENDING = 8
Const CREATE_IF_NOT_EXIST = True
Const FILENAME = "MYFILE.txt"

' --- Main Execution ---
' Get the current date and time
Dim nowtime
nowtime = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 3)

' Create the log message using a clear, multiline format
Dim logMessage
logMessage = "This program is written in VBScript." & vbCrLf & _
             "Current Time = " & nowtime & vbCrLf & _
             vbCrLf

' Call the subroutine to write the message to the file
Call AppendToFile(FILENAME, logMessage)

' --- Subroutines ---

' This subroutine handles the file writing process.
Sub AppendToFile(filePath, message)
    Dim fso, outfile
    
    ' Enable error handling
    On Error Resume Next
    
    ' Create the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open the file for appending, creating it if it doesn't exist.
    Set outfile = fso.OpenTextFile(filePath, FOR_APPENDING, CREATE_IF_NOT_EXIST)
    
    ' Check for errors during file opening
    If Err.Number <> 0 Then
        WScript.Echo "Error occurred while writing to file: " & Err.Description
        WScript.Quit
    End If
    
    ' Disable error handling
    On Error GoTo 0
    
    ' Write the message to the file
    outfile.WriteLine message
    WScript.Echo "Successfully appended message to " & filePath
    
    ' Clean up objects
    outfile.Close
    Set outfile = Nothing
    Set fso = Nothing
End Sub
