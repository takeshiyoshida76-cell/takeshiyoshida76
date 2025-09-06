package main

import (
    "fmt"
    "os"
    "time"
)

func main() {
    // ファイル名
    const filename = "MYFILE.txt"

    // 現在時刻をフォーマット
    nowtime := time.Now().Format("2006/01/02 15:04:05")

    // ファイルを追記モードで開く（存在しない場合は作成、書き込み専用）
    file, err := os.OpenFile(filename, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
    if err != nil {
        // エラーが発生した場合は、エラーメッセージを出力して終了
        fmt.Fprintf(os.Stderr, "ファイルを開けませんでした: %v\n", err)
        return
    }
    // 関数終了時にファイルを閉じることを保証
    defer file.Close()

    // ファイルにメッセージを書き込み
    _, err = fmt.Fprintln(file, "This program is written in Go.")
    if err != nil {
        fmt.Fprintf(os.Stderr, "ファイルへの書き込みに失敗しました: %v\n", err)
        return
    }

    _, err = fmt.Fprintf(file, "Current Time = %s\n", nowtime)
    if err != nil {
        fmt.Fprintf(os.Stderr, "ファイルへの書き込みに失敗しました: %v\n", err)
        return
    }
}
