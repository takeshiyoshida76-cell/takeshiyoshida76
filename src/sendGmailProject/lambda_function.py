# -*- coding: utf-8 -*-

import base64
import os
import json
import boto3
import datetime
import holidays
import requests
import google.auth
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
    if today in jp_holidays:
        print(f"今日は祝日 ({jp_holidays[today]}) なので、メールを送信しません。")
        return {
            'statusCode': 200,
            'body': json.dumps('今日は祝日のため、メール送信をスキップしました。')
        }

    try:
        # Gemini APIを呼び出して、要約とアクションを生成
        # エラーが発生した場合も、処理を継続するようにtry-exceptブロックで囲む
        try:
            generated_content = generate_summary_and_actions_with_gemini(report_text)
            # 要約・アクションのヘッダーを付加（空行を調整）
            generated_content_header = "\n\n要約と次へのアクション案(自動生成):\n"
        except Exception as e:
            print(f"Gemini API呼び出し中にエラーが発生しました。要約を付加せずにメールを送信します。エラー: {e}")
            generated_content = ""
            generated_content_header = ""

        # 文章組み立て
        email_body = report_text + generated_content_header + generated_content

        # 認証情報を取得してメール送信
        credentials = get_credentials()
        send_email(sender_email, recipient_email, email_subject, email_body, credentials, cc_email)
        
    except Exception as e:
        print(f"Lambda 関数エラー: {e}")
        return {
            'statusCode': 500,
            'body': json.dumps(f'メール送信エラー: {e}')
        }

    return {
        'statusCode': 200,
        'body': json.dumps('メール送信処理を実行しました。詳細は CloudWatch Logs を確認してください。')
    }

def generate_summary_and_actions_with_gemini(text):
    """
    Gemini APIを呼び出して、テキストの要約とアクションを生成します。
    """
    # 環境変数からAPIキーを取得
    api_key = os.environ.get('GEMINI_API_KEY')
    if not api_key:
        raise ValueError("環境変数 'GEMINI_API_KEY' が設定されていません。")

    # APIキーをクエリパラメータに追加
    #GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.0-pro:generateContent"
    #GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1/models/gemini-1.5-pro:generateContent"
    GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1/models/gemini-2.5-pro:generateContent"
    # Geminiへのプロンプトを一つのリクエストにまとめる
    prompt = f"""
    あなたは日報を分析するAIです。
    まず、次の5つの分析スタイルのうちどれか1つ、あるいは新しい視点を考え出してください。
    ・リスク検知（懸念点を洗い出し、対策を提案）
    ・成果最大化（成功事例を強調し、伸ばす方向を示す）
    ・チーム連携（人・役割のつながりに注目）
    ・戦略的思考（今後の展開・優先順位に焦点を当てる）
    ・自己改善（自分の成長や仕事の進め方の改善に焦点）
    ただし、上記に限らず、あなた自身で「今日はこの観点で読むと面白い」と思うテーマを自由に作ってもかまいません。
    次に、選んだ観点に沿って、以下のフォーマットで出力してください。余計な前置きや説明は一切入れないでください。

    出力形式:
    要約:
    [要約を2〜4行程度にまとめ、1行が長くなりすぎないよう適度に改行してください]

    テーマ:
    - [上記テーマを1行で完結に]

    アクション案:
    - [アクション1]
    - [アクション2]
    - [アクション3]
    
    日報テキスト:
    {text}
    """

    data = {
        "contents": [{"parts": [{"text": prompt}]}]
    }
    
    # APIキーを渡すためのヘッダーを作成
    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": api_key
    }
    response = requests.post(GEMINI_API_URL, headers=headers, json=data)
    response.raise_for_status() # HTTPエラーを発生させる
    
    result = response.json()
    
    if 'candidates' in result and len(result['candidates']) > 0:
        raw_output = result['candidates'][0]['content']['parts'][0]['text']
        # 余計な ``` や *** などのマークダウンを除去（念のため）
        cleaned = raw_output.strip().replace('```', '').replace('***', '')
        return cleaned
    else:
        print(f"APIからの応答に候補が含まれていませんでした。応答: {result}")
        return "APIからの応答がありませんでした。"

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
