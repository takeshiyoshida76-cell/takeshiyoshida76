Option Explicit

' AKASHI�Ζ��\���܂Ƃ߂ău���E�U�ŊJ���A���e��CSV�ɕۑ����鏈��
' ���s�O�ɁA�Ǘ��҂̃��[�U�ŁA��Ƀu���E�U���烍�O�C�����Ă�������

' --- �萔 ---
Const MURL = "https://atnd.ak4.jp/ja/manager/attendance/" ' AKASHI��URL
Const YYYYMM = "202509" ' �Ζ��\���擾�������N��
Const OUT_PATH = "c:\temp" ' �o�̓t�H���_
Const TEMP_FILE_NAME = "temp_akashi_temp.txt"
Const SAKURA_PATH = "C:\Program Files\sakura\sakura.exe" ' �� �T�N���G�f�B�^�̃p�X�����ɍ��킹�ĕύX���Ă������� ��

' --- �Ώێ҃��X�g ---
Dim g_persons
g_persons = Array("336546", "336415", "336530", "336540")
' �ѓc�A�v���A�����A�F��

' --- �I�u�W�F�N�g�̍쐬 ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_objShell: Set g_objShell = CreateObject("WScript.Shell")
Dim g_regex: Set g_regex = CreateObject("VBScript.RegExp")
g_regex.Pattern = "^\d{2}/\d{2}\([^\)]*\)" ' ���t�̃t�H�[�}�b�g

' --- ���C������ ---
Main()
Sub Main()
    InitializeEnvironment
    ProcessPersons
    WScript.Echo "���ׂĂ̏������������܂����B"
    CleanupEnvironment
End Sub

' --- ���ݒ� ---
Sub InitializeEnvironment()
    If Not g_fso.FolderExists(OUT_PATH) Then
        g_fso.CreateFolder OUT_PATH
    End If
    Dim tempFilePath: tempFilePath = OUT_PATH & "\" & TEMP_FILE_NAME
    If Not g_fso.FileExists(tempFilePath) Then
        Dim file: Set file = g_fso.CreateTextFile(tempFilePath, True)
        file.Close
    End If
End Sub

' --- �e�S���҂̏��� ---
Sub ProcessPersons()
    Dim person, url, outFile, tempFilePath

    For Each person In g_persons
        url = MURL & person & "/" & YYYYMM
        outFile = OUT_PATH & "\" & "�Ζ��\_" & person & "-" & YYYYMM & ".csv"
        tempFilePath = OUT_PATH & "\" & TEMP_FILE_NAME

        PerformAkashiExport person, url, tempFilePath
        ReadAndWriteData person, tempFilePath, outFile
    Next
End Sub

' --- AKASHI ����G�N�X�|�[�g ---
Sub PerformAkashiExport(person, url, tempFilePath)
    ' �u���E�U��URL���J��
    g_objShell.Run url
    ' �y�[�W�̃��[�h��ҋ@�i���߂ɐݒ�j
    WScript.Sleep 6000
    
    ' �S�I���ƃR�s�[
    g_objShell.SendKeys "^a"
    WScript.Sleep 1500 ' �S�I�������̂�ҋ@
    g_objShell.SendKeys "^c"
    ' �N���b�v�{�[�h�ւ̃R�s�[������ҋ@�i���߂ɐݒ�j
    WScript.Sleep 3500

    ' �T�N���G�f�B�^���N��
    g_objShell.Run """" & SAKURA_PATH & """ """ & tempFilePath & """", 0, False
    WScript.Sleep 3000
    
    ' �T�N���G�f�B�^�E�B���h�E���m���ɃA�N�e�B�u��
    g_objShell.AppActivate "sakura"
    WScript.Sleep 1500 ' �t�H�[�J�X���ڂ�̂�ҋ@
    
    ' �����̓��e��S�I�����č폜
    g_objShell.SendKeys "^a"
    WScript.Sleep 500
    g_objShell.SendKeys "{DEL}"
    WScript.Sleep 500
    
    ' �N���b�v�{�[�h�̓��e��\��t��
    g_objShell.SendKeys "^v"
    WScript.Sleep 2500
    
    ' �ۑ�����
    ' Ctrl+S�ŕۑ�
    g_objShell.SendKeys "^s"
    WScript.Sleep 1500
    
    ' �T�N���G�f�B�^�����
    g_objShell.SendKeys "%{F4}"
    WScript.Sleep 1000

    ' �u���E�U�E�B���h�E���A�N�e�B�u�����ă^�u�����
    ' �������s����ɂȂ�\�������邽�߁A���m���ȕ��@������
    ' �^�u�����
    g_objShell.SendKeys "^w" ' Ctrl+W�Ń^�u�����
    WScript.Sleep 1000
End Sub

' --- �ꎞ�t�@�C����ǂݍ��݁ACSV �t�@�C���ɏ����o�� ---
Sub ReadAndWriteData(person, tempFilePath, outFile)
    If g_fso.FileExists(tempFilePath) Then
        Dim objStream: Set objStream = CreateObject("ADODB.Stream")
        objStream.Charset = "UTF-8"
        objStream.Open
        objStream.LoadFromFile tempFilePath

        Dim objWriteFile: Set objWriteFile = g_fso.CreateTextFile(outFile, True, False)

        Dim strLine: strLine = ""
        Dim mergedLine1: mergedLine1 = ""
        Dim mergedLine2: mergedLine2 = ""
        Dim flgName: flgName = 0
        Dim char

        Do Until objStream.EOS
            char = objStream.ReadText(1)
            If char = vbCrLf Or char = vbLf Or char = vbCr Then
                strLine = Replace(strLine, vbTab, ",")
                ' �o�Ε�ƌ��x��\��
                If Left(strLine, 3) = "�o�Ε�" Or Mid(strLine, 8, 2) = "���x" Then
                    objWriteFile.WriteLine strLine
                End If
                ' ���O
                If flgName = 2 Then
                    objWriteFile.WriteLine strLine
                    flgName = 0
                End If
                If flgName = 1 Then
                    flgName = 2
                End If
                If Left(strLine, 7) = "�]�ƈ��I�����" Then
                    flgName = 1
                End If

                ' ���o���s
                If Left(strLine, 2) = "���t" Then
                    mergedLine1 = strLine & ","
                ElseIf mergedLine1 <> "" Then
                    mergedLine1 = mergedLine1 & strLine & ","
                    If Left(strLine, 5) = "�\��(��)" Then
                        mergedLine1 = Replace(mergedLine1, ",,", ",")
                        objWriteFile.WriteLine mergedLine1
                        mergedLine1 = ""
                    End If
                End If

                ' ���׍s
                If mergedLine2 <> "" Then
                    If g_regex.Test(strLine) Then
                        mergedLine2 = Replace(mergedLine2, ",,", ",")
                        mergedLine2 = Replace(mergedLine2, ",,", ",")
                        objWriteFile.WriteLine ReplaceSpecificComma(mergedLine2)
                        mergedLine2 = ""
                    Else
                        mergedLine2 = mergedLine2 & strLine & ","
                    End If
                End If
                If g_regex.Test(strLine) Then
                    mergedLine2 = strLine & ","
                End If
                strLine = ""
            Else
                strLine = strLine & char
            End If
        Loop
        ' �c��̖��׍s���o��
        If mergedLine2 <> "" Then
            mergedLine2 = Replace(mergedLine2, ",,", ",")
            mergedLine2 = Replace(mergedLine2, ",,", ",")
            objWriteFile.WriteLine ReplaceSpecificComma(mergedLine2)
        End If

        objStream.Close
        Set objStream = Nothing
        objWriteFile.Close
    Else
        WScript.Echo "�ꎞ�t�@�C�� " & tempFilePath & " ��������܂���B"
    End If
End Sub

' --- �Ζ��󋵗�����Z�߂ɂ��� ---
Function ReplaceSpecificComma(str)
    Dim parts, i
    parts = Split(str, ",") ' ������� "," �ŕ���
    
    ' 7�ڂ�8�ڂ̗v�f���`�F�b�N
    If UBound(parts) >= 8 Then
        Dim seventh, eighth
        seventh = parts(7) ' 7�ڂ̗v�f
        eighth = parts(8) ' 8�ڂ̗v�f
        If IsJapanese(seventh) And IsJapanese(eighth) Then
            ' 8�ڂ̒��O�� "," �� "�A" �ɕϊ�
            parts(7) = parts(7) & "�A" & parts(8)
            ' 8�ڂ��폜
            For i = 8 To UBound(parts) - 1
                parts(i) = parts(i + 1)
            Next
            ReDim Preserve parts(UBound(parts) - 1)
        End If
    End If
    
    ReplaceSpecificComma = Join(parts, ",") ' �C����̕������Ԃ�
End Function

' --- ���{��`�F�b�N ---
Function IsJapanese(str)
    If Len(str) > 0 Then
        Dim charCode
        charCode = AscW(Left(str, 1))
        ' ���� (U+4E00 - )
        If (charCode >= &H4E00) Then
            IsJapanese = True
        ' �Ђ炪�� (U+3040 - U+309F)
        ElseIf (charCode >= &H3040 And charCode <= &H309F) Then
            IsJapanese = True
        ' �J�^�J�i (U+30A0 - U+30FF)
        ElseIf (charCode >= &H30A0 And charCode <= &H30FF) Then
            IsJapanese = True
        Else
            IsJapanese = False
        End If
    Else
        IsJapanese = False
    End If
End Function

' --- ���̃N���[���A�b�v ---
Sub CleanupEnvironment()
    Set g_objShell = Nothing
    Set g_fso = Nothing
    Set g_regex = Nothing
End Sub
