import os
import random
import time
import msvcrt # Windows固有のキーボード入力を扱うライブラリ

# --- 定数と変数 ---
# ゲームボードの幅 (列数)
BOARD_WIDTH = 6
# ゲームボードの高さ (行数)
BOARD_HEIGHT = 12
# ぷよの種類 (色)
PUYO_CHARS = ["R", "G", "B", "Y"]

# ゲームボードを空のセル ("") で初期化
board = [["" for _ in range(BOARD_WIDTH)] for _ in range(BOARD_HEIGHT)]
# 現在操作中のぷよ
current_puyo = None
# 現在操作中のぷよのX座標
current_puyo_x = 0
# 現在操作中のぷよのY座標
current_puyo_y = 0

# --- ゲームの各機能 ---
def clear_screen():
    """
    画面をクリアする関数

    :詳細:
    Windows環境では 'cls' コマンド、macOS/Linux環境では 'clear' コマンドを使用して、
    コンソールの表示をクリアします。
    """
    os.system('cls' if os.name == 'nt' else 'clear')

def draw_board():
    """
    現在のゲームボードを画面に描画する関数

    :詳細:
    - 画面をクリアし、ゲームのタイトルを表示します。
    - ゲームボードの各セルをループでチェックし、ぷよや空のセル("・")を配置して表示します。
    - 現在操作中のぷよもボード上の正しい位置に描画します。
    - 最後に操作説明のテキストを表示します。
    """
    clear_screen()
    print("==================")
    print("  ぷよぷよ (Python)  ")
    print("==================")
    
    for y in range(BOARD_HEIGHT):
        row_str = ""
        for x in range(BOARD_WIDTH):
            if current_puyo and x == current_puyo_x and y == current_puyo_y:
                row_str += current_puyo + " "
            elif board[y][x] != "":
                row_str += board[y][x] + " "
            else:
                row_str += "・ "
        print(row_str)
    
    print("------------------")
    print("  ←: 左, →: 右, ↓: 落下, X: 終了")

def new_puyo():
    """
    新しいぷよを生成し、ボードの一番上に配置する関数

    :詳細:
    - グローバル変数`current_puyo`にランダムな色のぷよを割り当てます。
    - ぷよの初期位置をボード上部の中央に設定します。
    """
    global current_puyo, current_puyo_x, current_puyo_y
    current_puyo = random.choice(PUYO_CHARS)
    current_puyo_x = BOARD_WIDTH // 2
    current_puyo_y = 0

def check_and_clear_puyos():
    """
    4つ以上のぷよの塊をチェックし、消去する関数

    :詳細:
    - ボード全体を走査し、空でないセルから`find_cluster`関数を呼び出します。
    - 4つ以上の同じ色のぷよが繋がっているかを確認します。
    - ぷよの塊が見つかった場合、それらをボードから消去し、上にあるぷよを落下させます。
    - ぷよが消去された場合はTrueを、そうでない場合はFalseを返します。
    :return: ぷよが消去されたかどうか (bool)
    """
    cleared_puyos = set()
    
    for y in range(BOARD_HEIGHT):
        for x in range(BOARD_WIDTH):
            if board[y][x] != "":
                cluster = set()
                find_cluster(x, y, board[y][x], cluster)
                if len(cluster) >= 4:
                    cleared_puyos.update(cluster)
    
    if cleared_puyos:
        # 識別されたぷよを消去
        for x, y in cleared_puyos:
            board[y][x] = ""
        
        # ぷよを落下させる
        drop_puyos()
        return True
    return False

def find_cluster(x, y, color, cluster):
    """
    再帰的に同じ色の連結したぷよを見つける関数

    :param x: チェックを開始するぷよのX座標
    :param y: チェックを開始するぷよのY座標
    :param color: 探すぷよの色
    :param cluster: 見つかったぷよの座標を格納するセット
    :詳細:
    - ぷよのボード内での座標が有効であるか、色が一致するか、すでにクラスタに含まれていないかをチェックします。
    - 条件を満たす場合、そのぷよをクラスタに追加し、隣接する4方向のぷよに対してこの関数を再帰的に呼び出します。
    """
    if not (0 <= x < BOARD_WIDTH and 0 <= y < BOARD_HEIGHT):
        return
    if board[y][x] != color or (x, y) in cluster:
        return

    cluster.add((x, y))
    find_cluster(x + 1, y, color, cluster)
    find_cluster(x - 1, y, color, cluster)
    find_cluster(x, y + 1, color, cluster)
    find_cluster(x, y - 1, color, cluster)

def drop_puyos():
    """
    空のスペースを埋めるためにぷよを落下させる関数

    :詳細:
    - 各列を上から下へ走査します。
    - 空でないセルを見つけたら、そのぷよを列の一番下の空きスペースに移動させます。
    - これにより、ぷよが下へ自然に落下する効果を生み出します。
    """
    for x in range(BOARD_WIDTH):
        empty_y = BOARD_HEIGHT - 1
        for y in range(BOARD_HEIGHT - 1, -1, -1):
            if board[y][x] != "":
                board[empty_y][x] = board[y][x]
                if empty_y != y:
                    board[y][x] = ""
                empty_y -= 1

# --- メインゲームループ ---
def run_game():
    """
    ゲームを実行するメインの関数

    :詳細:
    - `new_puyo()`を呼び出して最初のぷよを生成します。
    - 無限ループに入り、以下の処理を繰り返します。
        - ユーザーのキー入力を処理し、ぷよを移動させます。
        - ぷよが落下可能かどうかをチェックし、可能であれば1マス下に移動させます。
        - 落下不可であれば、ぷよをボードに固定し、ゲームオーバーやぷよの消去判定を行います。
        - `X`キーが押された場合はループを抜けてゲームを終了します。
    """
    global current_puyo_x, current_puyo_y

    new_puyo()

    while True:
        draw_board()

        # Windows環境での非ブロッキングなキー入力を処理
        if msvcrt.kbhit():
            key = msvcrt.getch()
            
            # 矢印キーの処理
            if key == b'\xe0' or key == b'\x00':
                key = msvcrt.getch()
                if key == b'K' and current_puyo_x > 0: # 左矢印
                    current_puyo_x -= 1
                elif key == b'M' and current_puyo_x < BOARD_WIDTH - 1: # 右矢印
                    current_puyo_x += 1
                elif key == b'P': # 下矢印
                    # 一番下までぷよを落下させる
                    while current_puyo_y + 1 < BOARD_HEIGHT and board[current_puyo_y + 1][current_puyo_x] == "":
                        current_puyo_y += 1
            elif key.lower() == b'x':
                break # ゲームループを終了
        
        # 衝突判定
        if current_puyo_y + 1 >= BOARD_HEIGHT or board[current_puyo_y + 1][current_puyo_x] != "":
            # ぷよをボードに固定
            board[current_puyo_y][current_puyo_x] = current_puyo
            
            # ゲームオーバー判定 (ぷよが最上段に配置された場合)
            if current_puyo_y == 0:
                draw_board()
                print("==================")
                print("   ゲームオーバー   ")
                print("==================")
                break
                
            # 消去するぷよの塊をチェック
            while check_and_clear_puyos():
                draw_board()
                time.sleep(0.5)
            
            new_puyo()
        else:
            # ぷよを1マス下に移動
            current_puyo_y += 1
            
        time.sleep(0.2)

if __name__ == "__main__":
    run_game()
