Option Explicit

' AKASHI�Ζ��\��CSV���`�F�b�N���A�G���[���e��CSV�ɏo�͂��鏈��
' ���O�����ł��� akashiCsv2.vbs �����s�ςł��邱��

' --- �萔 ---
Const targetMonth = "202509" ' �Ζ��\���`�F�b�N�������N��
Const inputFolder = "c:\temp\" ' ���̓t�@�C���̃t�H���_

Dim outputFile
outputFile = inputFolder & "�Αӎw�E����_" & targetMonth & ".csv"  ' �G���[�t�@�C���̃t�@�C����

Dim regEx
Set regEx = New RegExp
regEx.Pattern = "^�Ζ��\_.*-" & targetMonth & "\.csv$"
regEx.IgnoreCase = True
regEx.Global = False

Dim fso, folder, output
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(inputFolder)
Set output = fso.CreateTextFile(outputFile, True)

output.WriteLine "�Ώێ�,���t,�G���[���R,�Ζ���,���R"

Dim file
For Each file In folder.Files
    If regEx.Test(file.Name) Then
        Call CheckFile(file.Path, output)
    End If
Next
WScript.Echo "�`�F�b�N�������������܂����B���ʂ� " & outputFile & _
             " �ɏo�͂���܂����B"

Sub CheckFile(filePath, output)
    Dim f, line, i, name, formatType, headerLine
    Set f = fso.OpenTextFile(filePath, 1)
    i = 0

    Do Until f.AtEndOfStream
        line = f.ReadLine
        i = i + 1

        If i = 3 Then
            name = line
        ElseIf i = 4 Then
            headerLine = line
            If InStr(headerLine, "�c�Ǝ���") > 0 Then
                formatType = "normal"
            Else
                formatType = "flex"
            End If
        ElseIf i > 4 Then
            Dim fields
            fields = Split(line, ",")

            If UBound(fields) < 15 Then Exit Sub

            Dim dateStr, clockIn, clockOut, actualIn, actualOut, plannedIn, plannedOut
            Dim status, totalTime, breakTime, overtime, nightTime, lateTime, reason
            Dim offset
            Dim rawReason
            Dim totalMin, breakMin, lateMin
            
            dateStr     = Trim(fields(0))   '���t
            clockIn     = Trim(fields(1))   '�ō�(�o)
            clockOut    = Trim(fields(2))   '�ō�(��)
            actualIn    = Trim(fields(3))   '����(�o)
            actualOut   = Trim(fields(4))   '����(��)
            plannedIn   = Trim(fields(5))   '�\��(�o)
            plannedOut  = Trim(fields(6))   '�\��(��)
            status      = Trim(fields(7))   '�Ζ���
            totalTime   = Trim(fields(8))   '���J������
            breakTime   = Trim(fields(9))   '�x�e����
            If formatType = "normal" Then
                offset = 0
            Else
                ' �t���b�N�X�͎c�Ǝ��ԗ񂪂Ȃ��̂ŁA�ȍ~�̗�1�O�|��
                offset = -1 
            End If
            nightTime   = Trim(fields(11 + offset)) '�[��J������
            lateTime    = Trim(fields(12 + offset)) '�x�����ގ���
            rawReason   = Trim(fields(15 + offset)) '���̗��R������
            
            ' ���R���̐擪�u���R �v���폜���郍�W�b�N
            If InStr(1, rawReason, "���R ") = 1 Then
                reason = Mid(rawReason, 4)
            Else
                reason = rawReason
            End If
            
            If formatType = "normal" Then
                overtime = Trim(fields(10)) '�c�Ǝ���
            Else
                If IsTime(totalTime) Then
                    totalMin = TimeToMinutes(totalTime)
                    ' ���J������-8:00���c�Ǝ��ԂƂ݂Ȃ�
                    overtime = MinutesToTime(totalMin - 480)
                Else
                    overtime = "0:00"
                End If
            End If

            ' �`�F�b�N1�F�ō��Y��{���R�s��
            If (IsTime(clockIn) = "" And IsTime(actualIn)) Or _
               (IsTime(clockOut) = "" And IsTime(actualOut)) Then
                If CheckReason(reason) = False Then
                    output.WriteLine name & "," & dateStr & _
                        ",�ō��Y�ꂪ���邪���R���L�ڂ���Ă��Ȃ�," & status & _
                        "," & reason
                End If
            End If

            ' �`�F�b�N3�F�x���{���R�s��
            If InStr(status, "�x��") > 0 Then
                If CheckReason(reason) = False Then
                    output.WriteLine name & "," & dateStr & _
                        ",�Ζ��󋵂ɒx�������邪���R���L�ڂ���Ă��Ȃ�," & status & _
                        "," & reason
                End If
            End If

            ' �`�F�b�N4�F�ދΎ��ԑO�ō��{���R�s��
            If IsTime(clockOut) And IsTime(plannedOut) Then
                If TimeToMinutes(clockOut) < TimeToMinutes(plannedOut) Then
                    If CheckReason(reason) = False Then
                        output.WriteLine name & "," & dateStr & _
                            ",�ދΗ\��O�ɑō�����Ă��邪���R���L�ڂ���Ă��Ȃ�," & _
                            status & "," & reason
                    End If
                End If
            End If

            ' �`�F�b�N5�F�ō��Ǝ��т̘����i30���ȏ�j
            If IsTime(clockIn) And IsTime(actualIn) Then
                If Abs(DateDiff("n", TimeValue(clockIn), _
                                TimeValue(actualIn))) >= 30 Then
                    If CheckReason(reason) = False Then
                        output.WriteLine name & "," & dateStr & _
                            ",�o�Αō��Ǝ��т�30���ȏ㘨�����Ă��邪�A���R���L�ڂ���Ă��Ȃ�," & _
                            status & "," & reason
                    End If
                End If
            End If

            ' �`�F�b�N9�F�ݑ�Ζ��L�ڂ�����̂ɋΖ��󋵂ɖ��L��
            If InStr(reason, "�ݑ�Ζ�") > 0 And InStr(status, "�ݑ�Ζ�") = 0 Then
                If InStr(reason, "�o��") = 0 And InStr(reason, "�o��") = 0 And _
                   InStr(reason, "�ړ�") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",���R���ɍݑ�Ζ��Ƃ��邪�Ζ��󋵂ɖ��L��," & status & _
                        "," & reason
                End if
            End If
            If InStr(status, "�ݑ�Ζ�") Then
                If InStr(reason, "�o��") > 0 Or InStr(reason, "�֏o��") > 0 Or _
                   InStr(reason, "�ɏo��") > 0 Or InStr(reason, "�ړ�") > 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",�Ζ��󋵂��ݑ�Ζ��Ƃ��邪���R���ɏo�Ђ�����," & status & _
                        "," & reason
                End if
            End If

            ' �`�F�b�N10�F�x�e���Ԃ�����Ȃ�
            If IsTime(totalTime) And IsTime(breakTime) Then
                totalMin = TimeToMinutes(totalTime)
                breakMin = TimeToMinutes(breakTime)
                If InStr(status, "�ߑO���N�x") Then
                    totalMin = totalMin - 180    '�ߑO���N�x�́A���J�����Ԃ���3���ԃ}�C�i�X
                ElseIf InStr(status, "�ߌ㔼�N�x") Then
                    totalMin = totalMin - 270    '�ߌ㔼�N�x�́A���J�����Ԃ���4.5���ԃ}�C�i�X
                ElseIf InStr(status, "�N�x") Or InStr(status, "�L�O���x��") Then
                    totalMin = totalMin - 450    '�N�x�E�L�O���x�ɂ́A���J�����Ԃ���7.5���ԃ}�C�i�X
                End If
                If totalMin > 360 And breakMin < 45 Then
                    output.WriteLine name & "," & dateStr & _
                        ",����6���Ԓ��ŋx�e45������," & status & "," & reason
                ElseIf totalMin > 480 And breakMin < 60 Then
                    output.WriteLine name & "," & dateStr & _
                        ",����8���Ԓ��ŋx�e1���Ԗ���," & status & "," & reason
                End If
            End If
            
            ' �`�F�b�N11�F�x�����R�̋L��
            If InStr(status, "�x��") > 0 And IsTime(lateTime) Then
                If TimeToMinutes(lateTime) > 0 Then
                    If CheckReason(reason) = False Then
                        output.WriteLine name & "," & dateStr & _
                            ",�x�����Ă��邪���R���ɒx�����R���L�ڂ���Ă��Ȃ�," & _
                            status & "," & reason
                    End If
                End If
            End If

            ' �`�F�b�N12�F�x�����̋x�e���ԏC��
            If InStr(status, "�x��") > 0 And IsTime(plannedIn) And IsTime(actualIn) Then
                ' �\��o�Ύ��ԂƎ��яo�Ύ��Ԃ̘�����3���Ԉȏ�ŁA�x�e���Ԃ�30���𒴂��Ă���ꍇ
                breakMin = TimeToMinutes(breakTime)
                lateMin = TimeToMinutes(lateTime)
                If lateMin >= 180 And breakMin > 30 Then
                    output.WriteLine name & "," & dateStr & _
                        ",�x�e���Ԃ��x�����ԂƂ��Ċ��Z����Ă���\������," & _
                        status & "," & reason
                End If
            End If

            ' �`�F�b�N13�F���ޗ��R�̋L��
            If InStr(status, "����") > 0 And IsTime(lateTime) Then
                If TimeToMinutes(lateTime) > 0 Then
                    If CheckReason(reason) = False Then
                        output.WriteLine name & "," & dateStr & _
                            ",���ނ��Ă��邪���R���ɑ��ޗ��R���L�ڂ���Ă��Ȃ�," & _
                            status & "," & reason
                    End If
                End If
            End If

            ' �`�F�b�N14�F���ގ��̋x�e���ԏC��
            If InStr(status, "����") > 0 And IsTime(plannedOut) And IsTime(actualOut) Then
                ' �\��o�Ύ��ԂƎ��яo�Ύ��Ԃ̘�����4���Ԉȏ�ŁA�x�e���Ԃ�30���𒴂��Ă���ꍇ
                breakMin = TimeToMinutes(breakTime)
                lateMin = TimeToMinutes(lateTime)
                If lateMin >= 240 And breakMin > 30 Then
                    output.WriteLine name & "," & dateStr & _
                        ",�x�e���Ԃ����ގ��ԂƂ��Ċ��Z����Ă���\������," & _
                        status & "," & reason
                End If
            End If
            
            ' �`�F�b�N15�F�d�Ԓx�����̏o�Ύ��ԂƋΖ���
            If InStr(status, "�d�Ԓx��") > 0 Then
                ' ���яo�Ύ��ԂƗ\��o�Ύ��Ԃ�����łȂ��`�F�b�N�i����̏ꍇ�G���[�j
                If clockIn = plannedIn Then
                    output.WriteLine name & "," & dateStr & _
                        ",�d�Ԓx�����R�����邪���яo�Ύ��Ԃ��\��Ɠ���ƂȂ��Ă���," & _
                        status & "," & reason
                End If
                ' �Ζ��󋵂Ɂu�x���v���܂܂�Ă���΃G���[
                If InStr(status, "�x��") > 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",�d�Ԓx���ɂ��ւ�炸�Ζ��󋵂ɒx�����܂܂�Ă���," & status & _
                        "," & reason
                End If
            End If

            ' �`�F�b�N16�F�c�Ɨ��R�Ǝw���҂��Ȃ�
            If IsTime(overtime) Then
                If TimeToMinutes(overtime) > 0 Then
                    If CheckReason(reason) = False Or InStr(reason, "�w��") = 0 Then
                        output.WriteLine name & "," & dateStr & _
                            ",�c�Ƃ����邪���R�܂��͎w���҂��L�ڂ���Ă��Ȃ�," & _
                            status & "," & reason
                    End If
                End If
            End If
            
            ' �`�F�b�N19�F�Ζ��󋵂��u�ߑO���N�x�v�܂��́u�ߌ㔼�N�x�v�̏ꍇ�A�c�Ƃ�����΃G���[�Ƃ���B
            If InStr(status, "�ߑO���N�x") > 0 Or InStr(status, "�ߌ㔼�N�x") > 0 Then
                If IsTime(overtime) Then
                    If TimeToMinutes(overtime) > 0 Then
                        output.WriteLine name & "," & dateStr & _
                            ",�����x�ɂɂ�������炸�c�Ƃ����͂���Ă���," & status & _
                            "," & reason
                    End If
                End If
            End If
            
            ' �`�F�b�N20�F�Ζ��󋵂��u�ߑO���N�x�v�܂��́u�ߌ㔼�N�x�v�̏ꍇ�A�Ζ����ԊO�̕s�v�ȋx�e���Ԃ����͂���Ă���΃G���[�Ƃ���B
            If InStr(status, "�ߑO���N�x") > 0 Or InStr(status, "�ߌ㔼�N�x") > 0 Then
                If IsTime(breakTime) Then
                    If TimeToMinutes(breakTime) >= 45 Then
                        ' �x�e���Ԃ�45���ȏ�̏ꍇ���G���[�Ƃ��܂��B
                        output.WriteLine name & "," & dateStr & _
                            ",�����x�ɂɂ�������炸�Ζ����ԊO�̕s�v�ȋx�e���Ԃ����͂���Ă���\������," & _
                            status & "," & reason
                    End If
                End If
            End If
            
            ' �`�F�b�N22�F�U�֏o�΂̏ꍇ�A���R���ɋx�o���R����юw���҂��L�ڂ���Ă��Ȃ���΃G���[�Ƃ���B
            If InStr(status, "�U�֏o��") > 0 Then
                If CheckReason(reason) = False Or InStr(reason, "�w��") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",�U�֏o�΂ɂ�������炸�A�x�o���R�܂��͎w���҂��L�ڂ���Ă��Ȃ�," & _
                        status & "," & reason
                End If
            End If
            
            ' �`�F�b�N23�F�U�֋x���̏ꍇ�A���R���Ɂu���������̐U�x�v�ƋL�ڂ��Ȃ���΃G���[�Ƃ���B
            If InStr(status, "�U�֋x��") > 0 Then
                If InStr(reason, "�̐U�x") = 0 And InStr(reason, "�̐U�֋x��") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",�U�֋x���ɂ�������炸�A�U�֌��̓��t���L�ڂ���Ă��Ȃ�," & _
                        status & "," & reason
                End If
            End If
            
            ' �`�F�b�N24�F�x���o�΂̏ꍇ�A���R���ɋx�o���R����юw���҂��L�ڂ���Ă��Ȃ���΃G���[�Ƃ���B
            If InStr(status, "�x���o��") > 0 Then
                If CheckReason(reason) = False Or InStr(reason, "�w��") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",�x���o�΂ɂ�������炸�A�x�o���R�܂��͎w���҂��L�ڂ���Ă��Ȃ�," & _
                        status & "," & reason
                End If
            End If
            
            ' �`�F�b�N26�F��x�̏ꍇ�A���R���Ɂu���������̑�x�v�ƋL�ڂ��Ȃ���΃G���[�Ƃ���B
            If InStr(status, "��x") > 0 Then
                If InStr(reason, "��x") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",��x�ɂ�������炸�A��x���̓��t���L�ڂ���Ă��Ȃ�," & status & _
                        "," & reason
                End If
            End If
            
            ' �`�F�b�N27�F�L�x�̏ꍇ�A�Ζ��󋵂ɍݑ�Ζ��������Ă���΃G���[�Ƃ���B
            If InStr(status, "�N�x") > 0 And InStr(status, "���N�x") = 0 Then
                If InStr(status, "�ݑ�Ζ�") > 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",�L�x�ƍݑ�Ζ��������ɋL�ڂ���Ă���," & status & _
                        "," & reason
                End If
            End If

        End If
    Loop
    f.Close
End Sub

Function IsTime(str)
    On Error Resume Next
    IsTime = IsDate("01/01/2000 " & str)
    On Error GoTo 0
End Function

Function TimeToMinutes(str)
    Dim t, result
    On Error Resume Next
    t = TimeValue(str)
    If Err.Number <> 0 Then
        result = 0
    Else
        result = Hour(t) * 60 + Minute(t)
    End If
    On Error GoTo 0
    TimeToMinutes = result
End Function

Function MinutesToTime(mins)
    Dim h, m
    h = Int(mins / 60)
    m = mins Mod 60
    MinutesToTime = h & ":" & Right("0" & m, 2)
End Function

Function CheckReason(str)
    If InStr(str, "���R") = 0 And InStr(str, "����") = 0 And _
       InStr(str, "��") = 0 And InStr(str, "��Â�") = 0 And _
       InStr(str, "��c") = 0 And InStr(str, "�ɂ��") = 0 Then
        CheckReason = False
    Else
        CheckReason = True
    End If
End Function
