#google_drive.py
#
# Google Drive 上のスプレッドシート・フォルダ・ファイルを一覧するためのヘルパー。
# 認証は google_auth_helper のトークン（drive.readonly スコープ）を流用する。

from datetime import datetime

import google_auth_helper

SPREADSHEET_MIME = "application/vnd.google-apps.spreadsheet"
FOLDER_MIME = "application/vnd.google-apps.folder"

# Drive のフォルダ／スプレッドシートを開くためのURL
FOLDER_URL = "https://drive.google.com/drive/folders/{id}"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/{id}/edit"


class DriveError(Exception):
    """Drive へのアクセスに失敗したことを表す（GUI 側でそのまま表示できる文面を持つ）"""


def _build_service():
    """Drive API のサービスを作る。未認可・未インストール時は DriveError を投げる。"""
    try:
        from googleapiclient.discovery import build
    except Exception as e:
        # 未インストールのほか、依存ライブラリが壊れている場合もここに来る
        raise DriveError(
            "Google Drive 連携に必要なライブラリを読み込めませんでした。\n"
            "pip install -r requirements.txt を実行してください。\n" + str(e)
        )

    if not google_auth_helper.has_token():
        raise DriveError(
            "Google の認可がまだ完了していません。\n"
            "「Google で認可する」ボタンから認可してください。"
        )

    try:
        creds = google_auth_helper.get_creds()
        if not creds:
            raise DriveError("トークンの取得に失敗しました。もう一度認可してください。")
        return build("drive", "v3", credentials=creds, cache_discovery=False)
    except DriveError:
        raise
    except Exception as e:
        raise DriveError(f"Google Drive への接続に失敗しました:\n{e}")


def _escape(text: str) -> str:
    """Drive の検索クエリ用に文字列をエスケープする。"""
    return str(text).replace("\\", "\\\\").replace("'", "\\'")


def _format_time(raw: str) -> str:
    """Drive の modifiedTime (RFC3339) を 'YYYY-MM-DD HH:MM' に整形する。"""
    if not raw:
        return ""
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).astimezone().strftime("%Y-%m-%d %H:%M")
    except Exception:
        return raw


def _list(query: str, limit: int = 200, fields: str = "files(id,name,modifiedTime,size,mimeType)") -> list:
    service = _build_service()
    items, page_token = [], None
    try:
        while True:
            response = service.files().list(
                q=query,
                pageSize=min(100, limit - len(items)),
                fields=f"nextPageToken,{fields}",
                orderBy="modifiedTime desc",
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
                pageToken=page_token,
            ).execute()
            items.extend(response.get("files", []))
            page_token = response.get("nextPageToken")
            if not page_token or len(items) >= limit:
                break
    except Exception as e:
        raise DriveError(f"Google Drive の一覧取得に失敗しました:\n{e}")
    return items


def list_spreadsheets(name_filter: str = "", limit: int = 200) -> list:
    """
    マイドライブ／共有ドライブ上のスプレッドシートを新しい順に返す。
    戻り値: [{"id", "name", "modified", "url"}]
    """
    query = f"mimeType='{SPREADSHEET_MIME}' and trashed=false"
    if name_filter.strip():
        query += f" and name contains '{_escape(name_filter.strip())}'"

    return [
        {
            "id": f["id"],
            "name": f.get("name", ""),
            "modified": _format_time(f.get("modifiedTime", "")),
            "url": SPREADSHEET_URL.format(id=f["id"]),
        }
        for f in _list(query, limit)
    ]


def list_folders(name_filter: str = "", limit: int = 200) -> list:
    """
    Drive 上のフォルダを新しい順に返す。
    戻り値: [{"id", "name", "modified", "url"}]
    """
    query = f"mimeType='{FOLDER_MIME}' and trashed=false"
    if name_filter.strip():
        query += f" and name contains '{_escape(name_filter.strip())}'"

    return [
        {
            "id": f["id"],
            "name": f.get("name", ""),
            "modified": _format_time(f.get("modifiedTime", "")),
            "url": FOLDER_URL.format(id=f["id"]),
        }
        for f in _list(query, limit)
    ]


def list_files_in_folder(folder_id: str, limit: int = 500) -> list:
    """
    指定フォルダ直下のファイル（フォルダ自身は除く）を返す。
    提出フォルダ監視で使うため、更新日時とサイズも含める。
    戻り値: [{"id", "name", "modified_epoch", "size"}]
    """
    query = f"'{_escape(folder_id)}' in parents and trashed=false and mimeType!='{FOLDER_MIME}'"

    files = []
    for f in _list(query, limit):
        raw = f.get("modifiedTime", "")
        try:
            epoch = datetime.fromisoformat(raw.replace("Z", "+00:00")).timestamp()
        except Exception:
            epoch = 0.0
        files.append({
            "id": f["id"],
            "name": f.get("name", ""),
            "modified_epoch": epoch,
            # Google ドキュメント形式のファイルには size が無いので 0 とする
            "size": int(f.get("size", 0) or 0),
        })
    return files


def extract_folder_id(text: str) -> str:
    """
    フォルダのURLまたはID文字列から、フォルダIDを取り出す。
    例: https://drive.google.com/drive/folders/ABC123?usp=sharing → ABC123
    """
    value = (text or "").strip()
    if not value:
        return ""
    if "/folders/" in value:
        value = value.split("/folders/", 1)[1]
    value = value.split("?", 1)[0].split("/", 1)[0]
    return value
