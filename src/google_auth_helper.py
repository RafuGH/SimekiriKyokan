#google_auth_helper.py

import os
import json
import webbrowser
import threading
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# スコープ
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# アプリデータディレクトリ
APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
os.makedirs(APP_DIR, exist_ok=True)

CREDENTIALS_PATH = os.path.join(APP_DIR, "credentials.json")  # ユーザーが置く
TOKEN_PATH = os.path.join(APP_DIR, "google_token.json")      # 自動生成


def get_creds():
    """
    保存されたトークンがあればそれを使用。
    なければ None を返す。
    """
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
        # トークン有効期限チェック・リフレッシュ
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            # 更新したトークンを保存
            with open(TOKEN_PATH, "w") as f:
                f.write(creds.to_json())
        return creds
    return None


def authorize_google_sheets(callback=None):
    """
    OAuth2 フロー を実行し、トークンを取得・保存する。
    callback: 認可完了時の処理（引数：success, message）
    戻り値: Credentials オブジェクト、または None
    """
    if not os.path.exists(CREDENTIALS_PATH):
        msg = f"credentials.json が見つかりません:\n{CREDENTIALS_PATH}\n\n" \
              "Google Cloud Console から credentials.json をダウンロードし、" \
              f"上記フォルダに配置してください。"
        if callback:
            callback(False, msg)
        return None

    try:
        flow = InstalledAppFlow.from_client_secrets_file(
            CREDENTIALS_PATH, SCOPES, 
            redirect_uri="http://localhost:8080/callback"
        )

        # ブラウザで認可ページを開く
        auth_url, state = flow.authorization_url()
        
        # ローカルサーバーを起動してコールバックを受ける
        def run_server():
            server = _CallbackServer(flow)
            try:
                server.handle_request()
            except Exception as e:
                print(f"Callback server error: {e}")

        server_thread = threading.Thread(target=run_server, daemon=True)
        server_thread.start()

        # ブラウザを自動起動
        webbrowser.open(auth_url)

        # サーバースレッドの終了を待つ（最大 30秒）
        server_thread.join(timeout=30)

        # トークンを取得
        creds = flow.credentials
        
        # トークンを保存
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())

        msg = "✅ Google Sheets の認可に成功しました！\n" \
              "トークンが保存されたので、次回からは URL だけで使用できます。"
        if callback:
            callback(True, msg)
        return creds

    except Exception as e:
        msg = f"❌ Google Sheets 認可に失敗しました:\n{str(e)}"
        if callback:
            callback(False, msg)
        return None


class _CallbackServer(BaseHTTPRequestHandler):
    """ローカルサーバーの認可コード受け取り用"""
    
    flow = None
    
    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)
        
        if "code" in params:
            # 認可コードを flow に渡す
            code = params["code"][0]
            try:
                self.flow.fetch_token(code=code)
                response = """
                <html>
                    <head><title>認可完了</title></head>
                    <body style="font-family: sans-serif; text-align: center; margin-top: 50px;">
                        <h2>✅ Google Sheets の認可が完了しました</h2>
                        <p>このウィンドウは閉じてください。</p>
                    </body>
                </html>
                """
                self.send_response(200)
            except Exception as e:
                response = f"""
                <html>
                    <head><title>エラー</title></head>
                    <body style="font-family: sans-serif; text-align: center; margin-top: 50px;">
                        <h2>❌ エラーが発生しました</h2>
                        <p>{str(e)}</p>
                    </body>
                </html>
                """
                self.send_response(400)
        else:
            response = """
            <html>
                <head><title>エラー</title></head>
                <body style="font-family: sans-serif; text-align: center; margin-top: 50px;">
                    <h2>❌ 認可コードを受け取れませんでした</h2>
                </body>
            </html>
            """
            self.send_response(400)

        self.send_header("Content-type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(response.encode("utf-8"))

    def log_message(self, format, *args):
        # サーバーログを無視
        pass


def has_token():
    """トークンが保存されているか確認"""
    return os.path.exists(TOKEN_PATH)


def clear_token():
    """保存されたトークンを削除"""
    if os.path.exists(TOKEN_PATH):
        os.remove(TOKEN_PATH)
