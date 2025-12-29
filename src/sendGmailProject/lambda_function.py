# -*- coding: utf-8 -*-

import base64
import os
import json
import boto3
import datetime
import holidays
import requests
import google.auth
import time
import random
import re
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.auth.exceptions import DefaultCredentialsError

# AWS Systems Manager Parameter Store クライアントを初期化
ssm_client = boto3.client('ssm')
# トークンを保存するParameter Storeのパラメータ名
TOKEN_PARAMETER_NAME = '/gmail/token' 

def lambda_handler(event, context):
    """
    Lambda 関数のエントリポイントです。
    """
    sender_email = os.environ.get('SENDER_EMAIL')
    recipient_email = os.environ.get('RECIPIENT_EMAILS')
    cc_email = os.environ.get('CC_EMAIL')
    your_name = os.environ.get('YOUR_NAME', '氏名')
    email_body_raw = os.environ.get('EMAIL_BODY')

    if not sender_email or not recipient_email or not email_body_raw:
        print("エラー: 環境変数 SENDER_EMAIL、RECIPIENT_EMAILS、EMAIL_BODY が設定されていません。")
        return {
            'statusCode': 500,
            'body': json.dumps('環境変数が正しく設定されていません。')
        }

    # デプロイ時に改行コードを正しく扱うための処理
    report_text = email_body_raw.replace('\\n', '\n')

    today = datetime.date.today()
    date_str = today.strftime("%Y/%m/%d")
    email_subject = f"【日報】{date_str}：{your_name}"

    jp_holidays = holidays.JP()
    if today in jp_holidays or is_year_end_holiday(today):
        reason = jp_holidays.get(today, "年末年始休暇")
        print(f"今日は休業日 ({reason}) なので、メールを送信しません。")
        return {
            'statusCode': 200,
            'body': json.dumps('今日は休業日のため、メール送信をスキップしました。')
        }

    # Gemini処理（例外時もメール継続）
    generated_content = ""
    generated_content_header = ""
    try:
        generated_content = generate_summary_and_actions_with_gemini(report_text)
        if generated_content:
            generated_content_header = "\n\n要約と次へのアクション案(自動生成):\n"
        else:
            print("Gemini API全失敗 → 要約なしでメール送信継続")
    except Exception as e:
        print(f"Gemini API呼び出し中にエラーが発生しました。要約を付加せずにメールを送信します。エラー: {e}")

    # メール本文組み立て（常に実行）
    email_body = report_text + generated_content_header + generated_content  # generated_contentは空でもOK

    # 認証/送信（例外時も継続）
    try:
        credentials = get_credentials()
        send_email(sender_email, recipient_email, email_subject, email_body, credentials, cc_email)
        print(f"メールを送信しました。宛先: {recipient_email}")
    except Exception as e:
        print(f"メール送信エラー: {e}")
        return {
            'statusCode': 500,
            'body': json.dumps(f'メール送信エラー: {e}')
        }

    return {
        'statusCode': 200,
        'body': json.dumps('メール送信処理を実行しました。詳細は CloudWatch Logs を確認してください。')
    }

def is_year_end_holiday(date):
    """
    12/30〜1/3 を年末年始休暇として判定
    """
    if date.month == 12 and date.day >= 30:
        return True
    if date.month == 1 and date.day <= 3:
        return True
    return False

def generate_summary_and_actions_with_gemini(text):
    api_key = os.environ.get('GEMINI_API_KEY')
    if not api_key:
        raise ValueError("GEMINI_API_KEY が未設定です")

    # 固有名詞をGeminiに送る前にマスキング
    masking_rules = load_masking_pairs()
    masked_text = apply_masking(text, masking_rules)
    masked_text = final_defense_masking(masked_text)

    #GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.0-pro:generateContent"
    #GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1/models/gemini-1.5-pro:generateContent"
    GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent"
    headers = {"Content-Type": "application/json", "x-goog-api-key": api_key}
    max_retries = 2
    base_delay = 30

    def call_gemini(prompt, step_name):
        data = {"contents": [{"parts": [{"text": prompt}]}]}
        for attempt in range(max_retries):
            try:
                response = requests.post(GEMINI_API_URL, headers=headers, json=data, timeout=30)

                if response.status_code == 200:
                    try:
                        result = response.json()
                        if 'candidates' in result and result['candidates']:
                            output = result['candidates'][0]['content']['parts'][0]['text'].strip()
                            print(f"【{step_name}成功】")
                            return output
                        else:
                            print(f"【{step_name}】候補なし")
                    except json.JSONDecodeError:
                        print(f"【{step_name}】JSONパース失敗")
                
                elif response.status_code >= 500:
                    error_msg = "詳細不明"
                    try:
                        error_msg = response.text[:100]
                    except:
                        pass
                    print(f"【{step_name}】{response.status_code}エラー: {error_msg}（試行{attempt+1}/{max_retries}）。リトライ...")
                
                elif response.status_code == 429:
                    retry_after = response.headers.get('Retry-After', base_delay)
                    try:
                        delay = int(retry_after)
                    except:
                        delay = base_delay
                    print(f"【{step_name}】429 Rate Limit（試行{attempt+1}/{max_retries}）。{delay}秒待機...")
                    time.sleep(delay)
                    continue
                
                else:
                    print(f"【{step_name}】{response.status_code}エラー: {response.text[:100]}")
                    return ""

            except requests.exceptions.Timeout:
                print(f"【{step_name}】タイムアウト（試行{attempt+1}/{max_retries}）。リトライ...")
            except requests.exceptions.ConnectionError as e:
                print(f"【{step_name}】接続エラー（試行{attempt+1}/{max_retries}）: {e}")
            except Exception as e:
                print(f"【{step_name}】予期せぬ例外（試行{attempt+1}/{max_retries}）: {e}")

            if attempt < max_retries - 1:
                try:
                    delay = base_delay + random.uniform(0, 10)
                    print(f"【{step_name}】{delay:.1f}秒待機...")
                    time.sleep(delay)
                except Exception as e:
                    print(f"【{step_name}】待機エラー: {e} → 30秒固定待機")
                    time.sleep(30)
        
        print(f"【{step_name}】全{max_retries}回リトライ失敗")
        return ""  # 失敗時は空文字列（メールに形跡なし）

    THEMES = [
        "リスク検知",
        "成果最大化",
        "チーム連携",
        "戦略的思考",
        "自己改善",
        "負荷分散",
        "属人性排除",
        "継続性設計",
        "意思決定の簡素化",
        "観測と可視化",
        "運用効率",
    ]

    # 段階1: 要約
    summary_prompt = f"""
    日報の要約を箇条書きで2〜4行でまとめてください。
    各行は1トピックのみ、40文字程度まで。
    文末は「。」で統一し、Markdown記号（*, -, •）は使わないでください。
    余計な前置きや説明は一切入れないでください。

    日報テキスト:
    {masked_text}
    """
    summary = call_gemini(summary_prompt, "要約生成")
    if not summary:
        return ""

    # 段階2: 視点
    theme = random.choice(THEMES)

    # 段階3: アクション
    action1_prompt = f"""
    以下の日報本文と視点に基づき、現実的で簡潔なアクション案を3つ提案してください。
    各アクションは1文・60文字以内を目安にしてください。
    背景説明や理由は不要です。

    出力は厳密に以下の形式で。

    - [アクション1]
    - [アクション2]
    - [アクション3]

    日報本文:
    {masked_text}

    視点:
    {theme}
    """
    actions = call_gemini(action1_prompt, "アクション生成") or ""  # 失敗時は空

    # 最終出力（空文字列耐性で安全、失敗文言なし）
    result = f"要約:\n{summary}\n\nテーマ:\n- {theme}\n\nアクション案:\n{actions}"

    # マスキングから復元
    result = restore_masking(result, masking_rules)

    return result if result.strip() else ""  # 空なら全体空

def load_masking_pairs():
    """
    環境変数 MASKING_PAIRS を読み込み、
    [(before, after), ...] のリストを返す
    """
    raw = os.environ.get("MASKING_PAIRS", "")
    pairs = []
    for item in raw.split("|"):
        if "=" in item:
            before, after = item.split("=", 1)
            pairs.append((before, after))
    return pairs

def apply_masking(text, pairs):
    """
    置換前 → 置換後
    """
    for before, after in pairs:
        text = text.replace(before, after)
    return text

def restore_masking(text, pairs):
    """
    置換後 → 置換前
    ※ 衝突防止のため長い文字列から戻す
    """
    for before, after in pairs:
        text = text.replace(after, before)
    return text

def final_defense_masking(text):
    """
    最終防衛マスキング（不可逆）
    ・会社名らしき表現をすべて「◎◎」で潰す
    """

    # ① アルファベット3文字以上（会社略称対策）
    text = re.sub(r'\b[A-Z]{3,}\b', '◎◎', text)

    # ② 「株式会社」の前側を2文字マスキング
    # 例: マイクロソフト株式会社 → マイクロソ◎◎株式会社
    text = re.sub(r'(.{2})株式会社', '◎◎株式会社', text)

    # ③ 「株式会社」の後側を2文字マスキング
    # 例: 株式会社Ｑ → 株式会社◎◎
    text = re.sub(r'株式会社(.{2})', '株式会社◎◎', text)

    # ④ 「社」の前側を2文字マスキング
    # 例: マイクロソフト社 → マイクロソ◎◎社
    text = re.sub(r'(.{2})社', '◎◎社', text)

    # ⑤ 「社」の後側を2文字マスキング
    # 例: Ｑ社 → ◎◎社
    text = re.sub(r'社(.{2})', '社◎◎', text)

    return text

def get_credentials():
    """
    Parameter Storeからトークンを取得し、必要であれば更新する
    """
    # Parameter Storeからトークンを取得
    try:
        response = ssm_client.get_parameter(
            Name=TOKEN_PARAMETER_NAME,
            WithDecryption=True # セキュアストリングとして保存した場合、復号化が必要
        )
        token_data = response['Parameter']['Value']
        # Parameter Storeから取得したトークンを/tmpに書き出す
        # from_authorized_user_fileはファイルパスを要求するため
        with open('/tmp/token.json', 'w') as f:
            f.write(token_data)
        creds = Credentials.from_authorized_user_file('/tmp/token.json', ['https://mail.google.com/'])
    except ssm_client.exceptions.ParameterNotFound:
        # トークンがParameter Storeに存在しない場合は、初回認証が必要です。
        raise Exception(f"Parameter Storeにトークン ({TOKEN_PARAMETER_NAME}) が見つかりません。初回認証が必要です。")
    except Exception as e:
        # その他のParameter Storeからのトークン取得エラー
        print(f"Parameter Storeからのトークン取得または読み込みエラー: {e}")
        raise

    # トークンが期限切れでリフレッシュトークンがある場合、更新する
    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            # 更新されたトークンをParameter Storeに保存
            ssm_client.put_parameter(
                Name=TOKEN_PARAMETER_NAME,
                Value=creds.to_json(),
                Type='SecureString', # セキュアストリングとして保存
                Overwrite=True # 既存のパラメータを上書き
            )
            # /tmpにも更新されたトークンを書き出す（次のAPI呼び出しのため）
            with open('/tmp/token.json', 'w') as token_file:
                token_file.write(creds.to_json())
            print("トークンをParameter Storeで更新しました。")
        except Exception as e:
            print(f"トークン更新エラー: {e}")
            raise 

    return creds

def send_email(sender, to, subject, body, credentials, cc=None):
    """
    メールを送信する
    """
    # MIMETextを作成
    msg = MIMEText(body, 'plain', 'utf-8')
    msg['to'] = to
    msg['from'] = sender
    msg['subject'] = subject
    if cc:
        msg['cc'] = cc
    
    # MIMEメッセージをBase64でエンコード
    message = {'raw': base64.urlsafe_b64encode(msg.as_bytes()).decode()}

    try:
        # Gmail APIのサービスを構築
        service = build('gmail', 'v1', credentials=credentials)
        # メールの送信
        service.users().messages().send(userId='me', body=message).execute()
        print(f"メールを送信しました。宛先: {to}")
    except Exception as e:
        print(f"メール送信エラー: {e}")
        raise
