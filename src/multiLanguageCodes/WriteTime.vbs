Const filename = "MYFILE.txt"

' Get system datetime & Edit message
Dim nowtime
nowtime = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 3)

Dim fso, outfile
Set fso = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
' Open File
Set outfile = fso.OpenTextFile(filename, 8, True) ' 8 = ForAppending, True = CreateIfNotExist
If Err.Number <> 0 Then
    WScript.Echo "ファイルへの書き込み中にエラーが発生しました: " & Err.Description
    WScript.Quit
End If
On Error GoTo 0

' Write Message
outfile.WriteLine "This program is written in VBScript."
outfile.WriteLine "Current Time = " & nowtime

' Close File
outfile.Close

Set outfile = Nothing
Set fso = Nothing
