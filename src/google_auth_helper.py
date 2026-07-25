#google_auth_helper.py

import os
import webbrowser
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# スコープ
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# アプリデータディレクトリ
APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
os.makedirs(APP_DIR, exist_ok=True)

CREDENTIALS_PATH = os.path.join(APP_DIR, "credentials.json")
TOKEN_PATH       = os.path.join(APP_DIR, "google_token.json")


def get_creds():
    """保存されたトークンがあればそれを使用。なければ None を返す。"""
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    if not os.path.exists(TOKEN_PATH):
        return None

    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
    return creds


def authorize_google_sheets(callback=None):
    """
    OAuth2 フローを別スレッドで実行し、トークンを取得・保存する。
    callback(success: bool, message: str) で結果を返す。
    """

    if not os.path.exists(CREDENTIALS_PATH):
        msg = (
            f"credentials.json が見つかりません:\n{CREDENTIALS_PATH}\n\n"
            "Google Cloud Console から credentials.json をダウンロードし、"
            "上記フォルダに配置してください。"
        )
        if callback:
            callback(False, msg)
        return

    def _run():
        try:
            from google_auth_oauthlib.flow import Flow

            flow = Flow.from_client_secrets_file(
                CREDENTIALS_PATH,
                scopes=SCOPES,
                redirect_uri="http://localhost:8080/callback"
            )

            auth_url, _ = flow.authorization_url(prompt="consent")

            # ブラウザを開く
            webbrowser.open(auth_url)

            # コールバック受け取り用
            result = {"code": None}
            done   = threading.Event()

            class _Handler(BaseHTTPRequestHandler):
                def do_GET(self):
                    parsed = urlparse(self.path)
                    params = parse_qs(parsed.query)
                    if "code" in params:
                        result["code"] = params["code"][0]
                        body = "<html><body><h2>✅ 認可完了。このウィンドウを閉じてください。</h2></body></html>"
                        self.send_response(200)
                    else:
                        body = "<html><body><h2>❌ エラーが発生しました。</h2></body></html>"
                        self.send_response(400)
                    self.send_header("Content-type", "text/html; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(body.encode("utf-8"))
                    done.set()

                def log_message(self, format, *args):
                    pass

            # サーバー起動（別スレッド）
            server = HTTPServer(("localhost", 8080), _Handler)
            server_thread = threading.Thread(target=server.handle_request, daemon=True)
            server_thread.start()

            # 最大120秒待機
            done.wait(timeout=120)
            server.server_close()

            if result["code"]:
                flow.fetch_token(code=result["code"])
                with open(TOKEN_PATH, "w") as f:
                    f.write(flow.credentials.to_json())
                if callback:
                    callback(True, "✅ Google Sheets の認可が完了しました！\n次回から自動でログインされます。")
            else:
                if callback:
                    callback(False, "❌ タイムアウトまたは認可がキャンセルされました。")

        except Exception as e:
            if callback:
                callback(False, f"❌ 認可に失敗しました:\n{str(e)}")

    # 完全に別スレッドで実行（GUI をブロックしない）
    threading.Thread(target=_run, daemon=True).start()


def has_token():
    """トークンが保存されているか確認"""
    return os.path.exists(TOKEN_PATH)


def clear_token():
    """保存されたトークンを削除"""
    if os.path.exists(TOKEN_PATH):
        os.remove(TOKEN_PATH)
