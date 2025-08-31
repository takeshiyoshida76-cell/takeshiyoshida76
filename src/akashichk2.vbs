Option Explicit

' AKASHI勤務表のCSVをチェックし、エラー内容をCSVに出力する処理
' 事前処理である akashiCsv2.vbs が実行済であること

' --- 定数 ---
Const targetMonth = "202509" ' 勤務表をチェックしたい年月
Const inputFolder = "c:\temp\" ' 入力ファイルのフォルダ

Dim outputFile
outputFile = inputFolder & "勤怠指摘事項_" & targetMonth & ".csv"  ' エラーファイルのファイル名

Dim regEx
Set regEx = New RegExp
regEx.Pattern = "^勤務表_.*-" & targetMonth & "\.csv$"
regEx.IgnoreCase = True
regEx.Global = False

Dim fso, folder, output
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(inputFolder)
Set output = fso.CreateTextFile(outputFile, True)

output.WriteLine "対象者,日付,エラー理由,勤務状況,理由"

Dim file
For Each file In folder.Files
    If regEx.Test(file.Name) Then
        Call CheckFile(file.Path, output)
    End If
Next
WScript.Echo "チェック処理が完了しました。結果は " & outputFile & _
             " に出力されました。"

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
            If InStr(headerLine, "残業時間") > 0 Then
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
            
            dateStr     = Trim(fields(0))   '日付
            clockIn     = Trim(fields(1))   '打刻(出)
            clockOut    = Trim(fields(2))   '打刻(退)
            actualIn    = Trim(fields(3))   '実績(出)
            actualOut   = Trim(fields(4))   '実績(退)
            plannedIn   = Trim(fields(5))   '予定(出)
            plannedOut  = Trim(fields(6))   '予定(退)
            status      = Trim(fields(7))   '勤務状況
            totalTime   = Trim(fields(8))   '総労働時間
            breakTime   = Trim(fields(9))   '休憩時間
            If formatType = "normal" Then
                offset = 0
            Else
                ' フレックスは残業時間列がないので、以降の列が1つ前倒し
                offset = -1 
            End If
            nightTime   = Trim(fields(11 + offset)) '深夜労働時間
            lateTime    = Trim(fields(12 + offset)) '遅刻早退時間
            rawReason   = Trim(fields(15 + offset)) '元の理由文字列
            
            ' 理由欄の先頭「理由 」を削除するロジック
            If InStr(1, rawReason, "理由 ") = 1 Then
                reason = Mid(rawReason, 4)
            Else
                reason = rawReason
            End If
            
            If formatType = "normal" Then
                overtime = Trim(fields(10)) '残業時間
            Else
                If IsTime(totalTime) Then
                    totalMin = TimeToMinutes(totalTime)
                    ' 総労働時間-8:00を残業時間とみなす
                    overtime = MinutesToTime(totalMin - 480)
                Else
                    overtime = "0:00"
                End If
            End If

            ' チェック1：打刻忘れ＋理由不備
            If (IsTime(clockIn) = "" And IsTime(actualIn)) Or _
               (IsTime(clockOut) = "" And IsTime(actualOut)) Then
                If CheckReason(reason) = False Then
                    output.WriteLine name & "," & dateStr & _
                        ",打刻忘れがあるが理由が記載されていない," & status & _
                        "," & reason
                End If
            End If

            ' チェック3：遅刻＋理由不備
            If InStr(status, "遅刻") > 0 Then
                If CheckReason(reason) = False Then
                    output.WriteLine name & "," & dateStr & _
                        ",勤務状況に遅刻があるが理由が記載されていない," & status & _
                        "," & reason
                End If
            End If

            ' チェック4：退勤時間前打刻＋理由不備
            If IsTime(clockOut) And IsTime(plannedOut) Then
                If TimeToMinutes(clockOut) < TimeToMinutes(plannedOut) Then
                    If CheckReason(reason) = False Then
                        output.WriteLine name & "," & dateStr & _
                            ",退勤予定前に打刻されているが理由が記載されていない," & _
                            status & "," & reason
                    End If
                End If
            End If

            ' チェック5：打刻と実績の乖離（30分以上）
            If IsTime(clockIn) And IsTime(actualIn) Then
                If Abs(DateDiff("n", TimeValue(clockIn), _
                                TimeValue(actualIn))) >= 30 Then
                    If CheckReason(reason) = False Then
                        output.WriteLine name & "," & dateStr & _
                            ",出勤打刻と実績が30分以上乖離しているが、理由が記載されていない," & _
                            status & "," & reason
                    End If
                End If
            End If

            ' チェック9：在宅勤務記載があるのに勤務状況に未記載
            If InStr(reason, "在宅勤務") > 0 And InStr(status, "在宅勤務") = 0 Then
                If InStr(reason, "出社") = 0 And InStr(reason, "出勤") = 0 And _
                   InStr(reason, "移動") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",理由欄に在宅勤務とあるが勤務状況に未記載," & status & _
                        "," & reason
                End if
            End If
            If InStr(status, "在宅勤務") Then
                If InStr(reason, "出社") > 0 Or InStr(reason, "へ出勤") > 0 Or _
                   InStr(reason, "に出勤") > 0 Or InStr(reason, "移動") > 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",勤務状況が在宅勤務とあるが理由欄に出社がある," & status & _
                        "," & reason
                End if
            End If

            ' チェック10：休憩時間が足りない
            If IsTime(totalTime) And IsTime(breakTime) Then
                totalMin = TimeToMinutes(totalTime)
                breakMin = TimeToMinutes(breakTime)
                If InStr(status, "午前半年休") Then
                    totalMin = totalMin - 180    '午前半年休は、総労働時間から3時間マイナス
                ElseIf InStr(status, "午後半年休") Then
                    totalMin = totalMin - 270    '午後半年休は、総労働時間から4.5時間マイナス
                ElseIf InStr(status, "年休") Or InStr(status, "記念日休暇") Then
                    totalMin = totalMin - 450    '年休・記念日休暇は、総労働時間から7.5時間マイナス
                End If
                If totalMin > 360 And breakMin < 45 Then
                    output.WriteLine name & "," & dateStr & _
                        ",実働6時間超で休憩45分未満," & status & "," & reason
                ElseIf totalMin > 480 And breakMin < 60 Then
                    output.WriteLine name & "," & dateStr & _
                        ",実働8時間超で休憩1時間未満," & status & "," & reason
                End If
            End If
            
            ' チェック11：遅刻理由の記載
            If InStr(status, "遅刻") > 0 And IsTime(lateTime) Then
                If TimeToMinutes(lateTime) > 0 Then
                    If CheckReason(reason) = False Then
                        output.WriteLine name & "," & dateStr & _
                            ",遅刻しているが理由欄に遅刻理由が記載されていない," & _
                            status & "," & reason
                    End If
                End If
            End If

            ' チェック12：遅刻時の休憩時間修正
            If InStr(status, "遅刻") > 0 And IsTime(plannedIn) And IsTime(actualIn) Then
                ' 予定出勤時間と実績出勤時間の乖離が3時間以上で、休憩時間が30分を超えている場合
                breakMin = TimeToMinutes(breakTime)
                lateMin = TimeToMinutes(lateTime)
                If lateMin >= 180 And breakMin > 30 Then
                    output.WriteLine name & "," & dateStr & _
                        ",休憩時間が遅刻時間として換算されている可能性あり," & _
                        status & "," & reason
                End If
            End If

            ' チェック13：早退理由の記載
            If InStr(status, "早退") > 0 And IsTime(lateTime) Then
                If TimeToMinutes(lateTime) > 0 Then
                    If CheckReason(reason) = False Then
                        output.WriteLine name & "," & dateStr & _
                            ",早退しているが理由欄に早退理由が記載されていない," & _
                            status & "," & reason
                    End If
                End If
            End If

            ' チェック14：早退時の休憩時間修正
            If InStr(status, "早退") > 0 And IsTime(plannedOut) And IsTime(actualOut) Then
                ' 予定出勤時間と実績出勤時間の乖離が4時間以上で、休憩時間が30分を超えている場合
                breakMin = TimeToMinutes(breakTime)
                lateMin = TimeToMinutes(lateTime)
                If lateMin >= 240 And breakMin > 30 Then
                    output.WriteLine name & "," & dateStr & _
                        ",休憩時間が早退時間として換算されている可能性あり," & _
                        status & "," & reason
                End If
            End If
            
            ' チェック15：電車遅延時の出勤時間と勤務状況
            If InStr(status, "電車遅延") > 0 Then
                ' 実績出勤時間と予定出勤時間が同一でないチェック（同一の場合エラー）
                If clockIn = plannedIn Then
                    output.WriteLine name & "," & dateStr & _
                        ",電車遅延理由があるが実績出勤時間が予定と同一となっている," & _
                        status & "," & reason
                End If
                ' 勤務状況に「遅刻」が含まれていればエラー
                If InStr(status, "遅刻") > 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",電車遅延にも関わらず勤務状況に遅刻が含まれている," & status & _
                        "," & reason
                End If
            End If

            ' チェック16：残業理由と指示者がない
            If IsTime(overtime) Then
                If TimeToMinutes(overtime) > 0 Then
                    If CheckReason(reason) = False Or InStr(reason, "指示") = 0 Then
                        output.WriteLine name & "," & dateStr & _
                            ",残業があるが理由または指示者が記載されていない," & _
                            status & "," & reason
                    End If
                End If
            End If
            
            ' チェック19：勤務状況が「午前半年休」または「午後半年休」の場合、残業があればエラーとする。
            If InStr(status, "午前半年休") > 0 Or InStr(status, "午後半年休") > 0 Then
                If IsTime(overtime) Then
                    If TimeToMinutes(overtime) > 0 Then
                        output.WriteLine name & "," & dateStr & _
                            ",半日休暇にもかかわらず残業が入力されている," & status & _
                            "," & reason
                    End If
                End If
            End If
            
            ' チェック20：勤務状況が「午前半年休」または「午後半年休」の場合、勤務時間外の不要な休憩時間が入力されていればエラーとする。
            If InStr(status, "午前半年休") > 0 Or InStr(status, "午後半年休") > 0 Then
                If IsTime(breakTime) Then
                    If TimeToMinutes(breakTime) >= 45 Then
                        ' 休憩時間が45分以上の場合をエラーとします。
                        output.WriteLine name & "," & dateStr & _
                            ",半日休暇にもかかわらず勤務時間外の不要な休憩時間が入力されている可能性あり," & _
                            status & "," & reason
                    End If
                End If
            End If
            
            ' チェック22：振替出勤の場合、理由欄に休出理由および指示者が記載されていなければエラーとする。
            If InStr(status, "振替出勤") > 0 Then
                If CheckReason(reason) = False Or InStr(reason, "指示") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",振替出勤にもかかわらず、休出理由または指示者が記載されていない," & _
                        status & "," & reason
                End If
            End If
            
            ' チェック23：振替休日の場合、理由欄に「●月●日の振休」と記載がなければエラーとする。
            If InStr(status, "振替休日") > 0 Then
                If InStr(reason, "の振休") = 0 And InStr(reason, "の振替休日") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",振替休日にもかかわらず、振替元の日付が記載されていない," & _
                        status & "," & reason
                End If
            End If
            
            ' チェック24：休日出勤の場合、理由欄に休出理由および指示者が記載されていなければエラーとする。
            If InStr(status, "休日出勤") > 0 Then
                If CheckReason(reason) = False Or InStr(reason, "指示") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",休日出勤にもかかわらず、休出理由または指示者が記載されていない," & _
                        status & "," & reason
                End If
            End If
            
            ' チェック26：代休の場合、理由欄に「●月●日の代休」と記載がなければエラーとする。
            If InStr(status, "代休") > 0 Then
                If InStr(reason, "代休") = 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",代休にもかかわらず、代休元の日付が記載されていない," & status & _
                        "," & reason
                End If
            End If
            
            ' チェック27：有休の場合、勤務状況に在宅勤務が入っていればエラーとする。
            If InStr(status, "年休") > 0 And InStr(status, "半年休") = 0 Then
                If InStr(status, "在宅勤務") > 0 Then
                    output.WriteLine name & "," & dateStr & _
                        ",有休と在宅勤務が同時に記載されている," & status & _
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
    If InStr(str, "理由") = 0 And InStr(str, "ため") = 0 And _
       InStr(str, "為") = 0 And InStr(str, "基づく") = 0 And _
       InStr(str, "会議") = 0 And InStr(str, "による") = 0 Then
        CheckReason = False
    Else
        CheckReason = True
    End If
End Function
