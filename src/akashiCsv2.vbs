Option Explicit

' AKASHI勤務表をまとめてブラウザで開き、内容をCSVに保存する処理
' 実行前に、管理者のユーザで、先にブラウザからログインしておくこと

' --- 定数 ---
Const MURL = "https://atnd.ak4.jp/ja/manager/attendance/" ' AKASHIのURL
Const YYYYMM = "202509" ' 勤務表を取得したい年月
Const OUT_PATH = "c:\temp" ' 出力フォルダ
Const TEMP_FILE_NAME = "temp_akashi_temp.txt"
Const SAKURA_PATH = "C:\Program Files\sakura\sakura.exe" ' ★ サクラエディタのパスを環境に合わせて変更してください ★

' --- 対象者リスト ---
Dim g_persons
g_persons = Array("336546", "336415", "336530", "336540")
' 葭田、久島、柴﨑、澁屋

' --- オブジェクトの作成 ---
Dim g_fso: Set g_fso = CreateObject("Scripting.FileSystemObject")
Dim g_objShell: Set g_objShell = CreateObject("WScript.Shell")
Dim g_regex: Set g_regex = CreateObject("VBScript.RegExp")
g_regex.Pattern = "^\d{2}/\d{2}\([^\)]*\)" ' 日付のフォーマット

' --- メイン処理 ---
Main()
Sub Main()
    InitializeEnvironment
    ProcessPersons
    WScript.Echo "すべての処理が完了しました。"
    CleanupEnvironment
End Sub

' --- 環境設定 ---
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

' --- 各担当者の処理 ---
Sub ProcessPersons()
    Dim person, url, outFile, tempFilePath

    For Each person In g_persons
        url = MURL & person & "/" & YYYYMM
        outFile = OUT_PATH & "\" & "勤務表_" & person & "-" & YYYYMM & ".csv"
        tempFilePath = OUT_PATH & "\" & TEMP_FILE_NAME

        PerformAkashiExport person, url, tempFilePath
        ReadAndWriteData person, tempFilePath, outFile
    Next
End Sub

' --- AKASHI からエクスポート ---
Sub PerformAkashiExport(person, url, tempFilePath)
    ' ブラウザでURLを開く
    g_objShell.Run url
    ' ページのロードを待機（長めに設定）
    WScript.Sleep 6000
    
    ' 全選択とコピー
    g_objShell.SendKeys "^a"
    WScript.Sleep 1500 ' 全選択されるのを待機
    g_objShell.SendKeys "^c"
    ' クリップボードへのコピー完了を待機（長めに設定）
    WScript.Sleep 3500

    ' サクラエディタを起動
    g_objShell.Run """" & SAKURA_PATH & """ """ & tempFilePath & """", 0, False
    WScript.Sleep 3000
    
    ' サクラエディタウィンドウを確実にアクティブ化
    g_objShell.AppActivate "sakura"
    WScript.Sleep 1500 ' フォーカスが移るのを待機
    
    ' 既存の内容を全選択して削除
    g_objShell.SendKeys "^a"
    WScript.Sleep 500
    g_objShell.SendKeys "{DEL}"
    WScript.Sleep 500
    
    ' クリップボードの内容を貼り付け
    g_objShell.SendKeys "^v"
    WScript.Sleep 2500
    
    ' 保存処理
    ' Ctrl+Sで保存
    g_objShell.SendKeys "^s"
    WScript.Sleep 1500
    
    ' サクラエディタを閉じる
    g_objShell.SendKeys "%{F4}"
    WScript.Sleep 1000

    ' ブラウザウィンドウをアクティブ化してタブを閉じる
    ' ここも不安定になる可能性があるため、より確実な方法を検討
    ' タブを閉じる
    g_objShell.SendKeys "^w" ' Ctrl+Wでタブを閉じる
    WScript.Sleep 1000
End Sub

' --- 一時ファイルを読み込み、CSV ファイルに書き出す ---
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
                ' 出勤簿と月度を表示
                If Left(strLine, 3) = "出勤簿" Or Mid(strLine, 8, 2) = "月度" Then
                    objWriteFile.WriteLine strLine
                End If
                ' 名前
                If flgName = 2 Then
                    objWriteFile.WriteLine strLine
                    flgName = 0
                End If
                If flgName = 1 Then
                    flgName = 2
                End If
                If Left(strLine, 7) = "従業員選択印刷" Then
                    flgName = 1
                End If

                ' 見出し行
                If Left(strLine, 2) = "日付" Then
                    mergedLine1 = strLine & ","
                ElseIf mergedLine1 <> "" Then
                    mergedLine1 = mergedLine1 & strLine & ","
                    If Left(strLine, 5) = "予定(退)" Then
                        mergedLine1 = Replace(mergedLine1, ",,", ",")
                        objWriteFile.WriteLine mergedLine1
                        mergedLine1 = ""
                    End If
                End If

                ' 明細行
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
        ' 残りの明細行を出力
        If mergedLine2 <> "" Then
            mergedLine2 = Replace(mergedLine2, ",,", ",")
            mergedLine2 = Replace(mergedLine2, ",,", ",")
            objWriteFile.WriteLine ReplaceSpecificComma(mergedLine2)
        End If

        objStream.Close
        Set objStream = Nothing
        objWriteFile.Close
    Else
        WScript.Echo "一時ファイル " & tempFilePath & " が見つかりません。"
    End If
End Sub

' --- 勤務状況欄を一纏めにする ---
Function ReplaceSpecificComma(str)
    Dim parts, i
    parts = Split(str, ",") ' 文字列を "," で分割
    
    ' 7個目と8個目の要素をチェック
    If UBound(parts) >= 8 Then
        Dim seventh, eighth
        seventh = parts(7) ' 7個目の要素
        eighth = parts(8) ' 8個目の要素
        If IsJapanese(seventh) And IsJapanese(eighth) Then
            ' 8個目の直前の "," を "、" に変換
            parts(7) = parts(7) & "、" & parts(8)
            ' 8個目を削除
            For i = 8 To UBound(parts) - 1
                parts(i) = parts(i + 1)
            Next
            ReDim Preserve parts(UBound(parts) - 1)
        End If
    End If
    
    ReplaceSpecificComma = Join(parts, ",") ' 修正後の文字列を返す
End Function

' --- 日本語チェック ---
Function IsJapanese(str)
    If Len(str) > 0 Then
        Dim charCode
        charCode = AscW(Left(str, 1))
        ' 漢字 (U+4E00 - )
        If (charCode >= &H4E00) Then
            IsJapanese = True
        ' ひらがな (U+3040 - U+309F)
        ElseIf (charCode >= &H3040 And charCode <= &H309F) Then
            IsJapanese = True
        ' カタカナ (U+30A0 - U+30FF)
        ElseIf (charCode >= &H30A0 And charCode <= &H30FF) Then
            IsJapanese = True
        Else
            IsJapanese = False
        End If
    Else
        IsJapanese = False
    End If
End Function

' --- 環境のクリーンアップ ---
Sub CleanupEnvironment()
    Set g_objShell = Nothing
    Set g_fso = Nothing
    Set g_regex = Nothing
End Sub
