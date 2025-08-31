import os
import time
import copy

# === ゲーム設定 ===
BOARD_SIZE = 8
PLAYER_STONE = "@"  # プレイヤーの石（黒）
COMPUTER_STONE = "O"  # コンピュータの石（白）
EMPTY = "."  # 空のセル
DIRECTIONS = [(0, 1), (0, -1), (1, 0), (-1, 0), (1, 1), (1, -1), (-1, 1), (-1, -1)]  # 8方向
CORNERS = [(0, 0), (0, 7), (7, 0), (7, 7)]  # ボードの角
MINIMAX_DEPTH = 5  # ミニマックスの探索深さ
CORNER_VALUE = 25  # 角の評価点
MOBILITY_VALUE = 5  # 着手可能数の評価点

# 初期ボード（中央4マスに黒白の石を配置）
INITIAL_BOARD = [[EMPTY] * BOARD_SIZE for _ in range(BOARD_SIZE)]
INITIAL_BOARD[3][3], INITIAL_BOARD[3][4] = PLAYER_STONE, COMPUTER_STONE
INITIAL_BOARD[4][3], INITIAL_BOARD[4][4] = COMPUTER_STONE, PLAYER_STONE

# === 画面表示 ===
def clear_screen():
    """
    コンソール画面をクリアする。

    処理:
        1. OSに応じて適切なクリアコマンドを実行（Windowsは'cls'、他は'clear'）。

    戻り値:
        なし
    """
    os.system('cls' if os.name == 'nt' else 'clear')

def draw_board(board):
    """
    ゲームボードをコンソールに表示する。

    引数:
        board: 8x8の2次元リスト（ボード状態）

    処理:
        1. コンソールをクリア。
        2. 列ラベル（a-h）を表示。
        3. 各行を番号（1-8）とともに表示。
        4. ボードの罫線を表示。

    戻り値:
        なし
    """
    clear_screen()
    print("  a b c d e f g h")  # 列ラベル
    print(" " + "-" * 17)  # 上部罫線
    for i, row in enumerate(board, 1):
        print(f"{i}|{' '.join(row)}")  # 行番号とボード内容
    print(" " + "-" * 17)  # 下部罫線

def display_game_state(board):
    """
    現在のボードと石の数を表示する。

    引数:
        board: 8x8の2次元リスト（ボード状態）

    処理:
        1. ボードを表示。
        2. プレイヤーとコンピュータの石の数をカウント。
        3. 石の数を表示。

    戻り値:
        なし
    """
    draw_board(board)  # ボード表示
    player_count, computer_count = get_stone_counts(board)  # 石の数取得
    print(f"黒(あなた): {player_count} 白(コンピュータ): {computer_count}")

def display_end_game_results(board):
    """
    ゲーム終了時のボードと結果を表示する。

    引数:
        board: 8x8の2次元リスト（ボード状態）

    処理:
        1. 現在のボードと石の数を表示。
        2. 最終結果（石の数と勝敗）を表示。
        3. ユーザーのキー入力を待つ。

    戻り値:
        なし
    """
    display_game_state(board)  # ボードと石の数表示
    print("=" * 20)
    print("最終結果")
    player_count, computer_count = get_stone_counts(board)  # 最終石数
    print(f"黒(あなた): {player_count} 白(コンピュータ): {computer_count}")
    if player_count > computer_count:
        print("あなたの勝ち！")
    elif player_count < computer_count:
        print("コンピュータの勝ち。")
    else:
        print("引き分け。")
    print("=" * 20)
    input("終了するにはキーを押してください...")  # 終了待機

# === ゲームロジック ===
def get_stone_counts(board):
    """
    ボード上のプレイヤーとコンピュータの石の数をカウントする。

    引数:
        board: 8x8の2次元リスト（ボード状態）

    処理:
        1. 各行を走査し、プレイヤーとコンピュータの石をカウント。
        2. カウント結果をタプルで返す。

    戻り値:
        tuple: (プレイヤーの石の数, コンピュータの石の数)
    """
    player_count = sum(row.count(PLAYER_STONE) for row in board)
    computer_count = sum(row.count(COMPUTER_STONE) for row in board)
    return player_count, computer_count

def is_valid_move(board, row, col, stone):
    """
    指定位置への石の配置が有効かを判定する。

    引数:
        board: 8x8の2次元リスト（ボード状態）
        row: 行インデックス（0-7）
        col: 列インデックス（0-7）
        stone: 配置する石（PLAYER_STONEまたはCOMPUTER_STONE）

    処理:
        1. 指定位置が空でない場合、無効と判定。
        2. 8方向をチェックし、相手の石を挟めるか確認。
        3. 挟める場合、Trueを返す。

    戻り値:
        bool: 配置が有効ならTrue、さもなくばFalse
    """
    if board[row][col] != EMPTY:
        return False
    opponent_stone = COMPUTER_STONE if stone == PLAYER_STONE else PLAYER_STONE
    for dr, dc in DIRECTIONS:
        r, c = row + dr, col + dc
        if 0 <= r < BOARD_SIZE and 0 <= c < BOARD_SIZE and board[r][c] == opponent_stone:
            r, c = r + dr, c + dc
            while 0 <= r < BOARD_SIZE and 0 <= c < BOARD_SIZE:
                if board[r][c] == stone:
                    return True  # 挟める石が見つかった
                if board[r][c] == EMPTY:
                    break  # 空マスで終了
                r, c = r + dr, c + dc
    return False

def get_valid_moves(board, stone):
    """
    指定された石の有効な手をすべて取得する。

    引数:
        board: 8x8の2次元リスト（ボード状態）
        stone: 対象の石（PLAYER_STONEまたはCOMPUTER_STONE）

    処理:
        1. ボード全体を走査。
        2. 空マスで有効な手をチェック。
        3. 有効な手の座標リストを返す。

    戻り値:
        list: 有効な手の座標リスト（(row, col)のタプル）
    """
    return [(i, j) for i in range(BOARD_SIZE) for j in range(BOARD_SIZE)
            if board[i][j] == EMPTY and is_valid_move(board, i, j, stone)]

def flip_stones(board, row, col, stone):
    """
    石を配置し、挟まれた相手の石を反転する。

    引数:
        board: 8x8の2次元リスト（ボード状態）
        row: 行インデックス（0-7）
        col: 列インデックス（0-7）
        stone: 配置する石（PLAYER_STONEまたはCOMPUTER_STONE）

    処理:
        1. 指定位置に石を配置。
        2. 8方向をチェックし、挟める相手の石を特定。
        3. 挟める石を反転。

    戻り値:
        なし（ボードを直接更新）
    """
    board[row][col] = stone  # 石を配置
    opponent_stone = COMPUTER_STONE if stone == PLAYER_STONE else PLAYER_STONE
    for dr, dc in DIRECTIONS:
        r, c = row + dr, col + dc
        stones_to_flip = []
        while 0 <= r < BOARD_SIZE and 0 <= c < BOARD_SIZE and board[r][c] == opponent_stone:
            stones_to_flip.append((r, c))  # 反転候補を記録
            r, c = r + dr, c + dc
        if 0 <= r < BOARD_SIZE and 0 <= c < BOARD_SIZE and board[r][c] == stone:
            for fr, fc in stones_to_flip:
                board[fr][fc] = stone  # 挟める石を反転

def check_game_end(board, player_moves, computer_moves, pass_count):
    """
    ゲームの終了条件を判定する。

    引数:
        board: 8x8の2次元リスト（ボード状態）
        player_moves: プレイヤーの有効な手のリスト
        computer_moves: コンピュータの有効な手のリスト
        pass_count: 連続パスの回数

    処理:
        1. 両者が有効な手を持たない場合、ゲーム終了。
        2. 連続パスが2回以上の場合、ゲーム終了。
        3. 終了時にメッセージを表示。

    戻り値:
        bool: ゲーム終了ならTrue、さもなくばFalse
    """
    if not player_moves and not computer_moves:
        print("両者とも有効な手がありません。ゲーム終了。")
        return True
    if pass_count >= 2:
        print("連続パスによりゲーム終了。")
        return True
    return False

# === AIロジック ===
def evaluate_board(board):
    """
    ボードを評価してスコアを計算する。

    引数:
        board: 8x8の2次元リスト（ボード状態）

    処理:
        1. 角の石を評価（コンピュータ有利なら加点、プレイヤー有利なら減点）。
        2. 着手可能数を評価（コンピュータが多いと加点）。
        3. 石の数を評価（コンピュータが多いと加点）。
        4. 総合スコアを返す。

    戻り値:
        int: ボードの評価スコア（コンピュータにとって高いほど良い）
    """
    score = 0
    player_count, computer_count = get_stone_counts(board)  # 石の数取得
    # 角の評価
    for r, c in CORNERS:
        if board[r][c] == COMPUTER_STONE:
            score += CORNER_VALUE  # コンピュータが角を確保
        elif board[r][c] == PLAYER_STONE:
            score -= CORNER_VALUE  # プレイヤーが角を確保
    # 着手可能数の評価
    computer_mobility = len(get_valid_moves(board, COMPUTER_STONE))
    player_mobility = len(get_valid_moves(board, PLAYER_STONE))
    score += (computer_mobility - player_mobility) * MOBILITY_VALUE
    # 石の数の評価
    score += computer_count - player_count
    return score

def minimax(board, depth, is_maximizing, alpha, beta):
    """
    アルファベータ枝刈り付きミニマックスでボードの最善手を評価する。

    引数:
        board: 8x8の2次元リスト（ボード状態）
        depth: 残りの探索深さ（整数）
        is_maximizing: コンピュータ（最大化: True）かプレイヤー（最小化: False）か
        alpha: 最大化プレイヤーが確保できる最低スコア（整数またはfloat('-inf')）
        beta: 最小化プレイヤーが許容する最高スコア（整数またはfloat('inf')）

    処理:
        1. 終了条件（深さ0またはゲーム終了）ならボードを評価してスコアを返す。
        2. 最大化（コンピュータ）の場合、可能な手の中で最高スコアを選択し、αを更新。
        3. 最小化（プレイヤー）の場合、可能な手の中で最低スコアを選択し、βを更新。
        4. α ≥ β なら枝刈りして探索を終了。
        5. 有効な手がない場合はパスして相手のターンで再評価。

    戻り値:
        int: 評価スコア（コンピュータにとって高いほど良い）
    """
    # 終了条件: 深さが0またはゲーム終了（両者とも手がない）
    if depth == 0 or not (get_valid_moves(board, PLAYER_STONE) or get_valid_moves(board, COMPUTER_STONE)):
        return evaluate_board(board)

    # コンピュータのターン（スコアを最大化）
    if is_maximizing:
        best_score = float('-inf')
        valid_moves = get_valid_moves(board, COMPUTER_STONE)
        if not valid_moves:
            # 有効な手がない場合、パスして相手のターンで再評価
            return minimax(board, depth - 1, False, alpha, beta)
        for move in valid_moves:
            # 手を試し、次の状態を評価
            new_board = copy.deepcopy(board)
            flip_stones(new_board, move[0], move[1], COMPUTER_STONE)
            score = minimax(new_board, depth - 1, False, alpha, beta)
            best_score = max(best_score, score)
            alpha = max(alpha, best_score)  # アルファを更新
            if beta <= alpha:
                break  # 枝刈り: これ以上の探索は不要
        return best_score
    
    # プレイヤーのターン（スコアを最小化）
    else:
        best_score = float('inf')
        valid_moves = get_valid_moves(board, PLAYER_STONE)
        if not valid_moves:
            # 有効な手がない場合、パスして相手のターンで再評価
            return minimax(board, depth - 1, True, alpha, beta)
        for move in valid_moves:
            # 手を試し、次の状態を評価
            new_board = copy.deepcopy(board)
            flip_stones(new_board, move[0], move[1], PLAYER_STONE)
            score = minimax(new_board, depth - 1, True, alpha, beta)
            best_score = min(best_score, score)
            beta = min(beta, best_score)  # ベータを更新
            if beta <= alpha:
                break  # 枝刈り: これ以上の探索は不要
        return best_score

def find_best_move(board, stone, depth):
    """
    アルファベータ枝刈り付きミニマックスで最善手を選択する。

    引数:
        board: 8x8の2次元リスト（ボード状態）
        stone: 対象の石（PLAYER_STONEまたはCOMPUTER_STONE）
        depth: ミニマックスの探索深さ（整数）

    処理:
        1. 有効な手をすべて取得。
        2. 各手を試し、アルファベータ枝刈り付きミニマックスでスコアを評価。
        3. 最高スコアの手を選択。
        4. 有効な手がない場合、Noneを返す。

    戻り値:
        tuple or None: 最善手の座標（row, col）またはNone（有効な手がない場合）
    """
    valid_moves = get_valid_moves(board, stone)
    if not valid_moves:
        return None
    best_score = float('-inf')
    best_move = None
    alpha = float('-inf')  # 初期アルファ
    beta = float('inf')   # 初期ベータ
    for move in valid_moves:
        new_board = copy.deepcopy(board)
        flip_stones(new_board, move[0], move[1], stone)  # 手を試す
        score = minimax(new_board, depth - 1, False, alpha, beta)  # 次の状態を評価
        if score > best_score:
            best_score = score
            best_move = move
        alpha = max(alpha, best_score)  # アルファを更新
    return best_move

# === ターン処理 ===
def handle_player_turn(board):
    """
    プレイヤーのターンを処理する。

    引数:
        board: 8x8の2次元リスト（ボード状態）

    処理:
        1. 有効な手を確認。
        2. 有効な手がない場合、パスを通知しFalseを返す。
        3. ユーザーに入力を求め、有効性を検証。
        4. 'exit'が入力された場合、Noneを返す。
        5. 有効な手が入力された場合、石を配置しTrueを返す。

    戻り値:
        bool or None: 手が置けた場合はTrue、パスならFalse、終了ならNone
    """
    valid_moves = get_valid_moves(board, PLAYER_STONE)
    if not valid_moves:
        print("黒: 有効な手がありません。パスします。")
        time.sleep(1)  # メッセージを読みやすくする待機
        return False
    print("あなたのターン（黒）。")
    while True:
        move_str = input("配置場所（例: d3, c5）または 'exit' で終了: ").strip().lower()
        if move_str == "exit":
            print("ゲームを終了します。")
            time.sleep(1)  # メッセージを読みやすくする待機
            return None
        if len(move_str) != 2 or move_str[0] not in 'abcdefgh' or move_str[1] not in '12345678':
            print("無効な入力です（例: a1）。")
            continue
        col, row = ord(move_str[0]) - ord('a'), int(move_str[1]) - 1
        if (row, col) in valid_moves:
            flip_stones(board, row, col, PLAYER_STONE)  # 石を配置
            return True
        print("その場所には置けません。")

def handle_computer_turn(board):
    """
    コンピュータのターンを処理する。

    引数:
        board: 8x8の2次元リスト（ボード状態）

    処理:
        1. 有効な手を確認。
        2. 有効な手がない場合、パスを通知しFalseを返す。
        3. ミニマックスで最善手を選択。
        4. 選択した手に石を配置し、Trueを返す。

    戻り値:
        bool: 手が置けた場合はTrue、パスならFalse
    """
    valid_moves = get_valid_moves(board, COMPUTER_STONE)
    if not valid_moves:
        print("白: 有効な手がありません。パスします。")
        time.sleep(1)  # メッセージを読みやすくする待機
        return False
    print("コンピュータのターン（白）。")
    time.sleep(1)  # 思考を演出（2秒から1秒に短縮）
    move = find_best_move(board, COMPUTER_STONE, MINIMAX_DEPTH)
    if move:
        row, col = move
        print(f"コンピュータが {chr(ord('a') + col)}{row + 1} に置きました。")
        time.sleep(1)  # メッセージを読みやすくする待機
        flip_stones(board, row, col, COMPUTER_STONE)  # 石を配置
    return True

# === メインループ ===
def main():
    """
    オセロゲームのメインループを実行する。

    処理:
        1. 初期ボードを準備。
        2. プレイヤーとコンピュータのターンを交互に処理。
        3. ゲーム終了条件を満たすまでループ。
        4. 終了時に結果を表示。

    戻り値:
        なし
    """
    board = copy.deepcopy(INITIAL_BOARD)
    current_turn = PLAYER_STONE
    pass_count = 0

    while True:
        display_game_state(board)  # ボード表示
        player_moves = get_valid_moves(board, PLAYER_STONE)
        computer_moves = get_valid_moves(board, COMPUTER_STONE)

        if check_game_end(board, player_moves, computer_moves, pass_count):
            display_end_game_results(board)  # 最終結果表示
            return

        if current_turn == PLAYER_STONE:
            result = handle_player_turn(board)  # プレイヤーのターン
            if result is None:  # 終了選択
                return
            pass_count = 0 if result else pass_count + 1
        else:
            result = handle_computer_turn(board)  # コンピュータのターン
            pass_count = 0 if result else pass_count + 1

        current_turn = COMPUTER_STONE if current_turn == PLAYER_STONE else PLAYER_STONE  # ターン交代

if __name__ == "__main__":
    main()