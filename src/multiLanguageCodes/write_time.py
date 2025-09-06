# -*- coding: utf-8 -*-
from datetime import datetime

filename = "MYFILE.txt"

def write_time_program():
    # Get system datetime & edit message
    now = datetime.now()
    nowtime_str = now.strftime("%Y/%m/%d %H:%M:%S")

    # Open file
    try:
        with open(filename, "a", encoding="utf-8") as f:
            # TALプログラムであることを示すメッセージを書き込む
            f.write("This program is written in Python.\n")
            # 現在時刻を書き込む
            f.write(f"Current Time = {nowtime_str}\n")
    except IOError as e:
        print(f"ファイルへの書き込み中にエラーが発生しました: {e}")

if __name__ == "__main__":
    write_time_program()
    print(f"ログが'{filename}'に追記されました。")
