#submission_watcher.py
#
# 提出フォルダ（OneDrive / SharePoint の同期フォルダなど）を監視し、
# 新しく置かれた・更新されたファイルを検出して Webhook 通知する。
#
# OneDrive や Teams の「ファイル」タブは、PC上では通常のフォルダとして同期されるため、
# その同期フォルダを監視することで「チームへの提出」を検知できる。

import json
import os
from datetime import datetime

from webhook_client import send_webhook_text

APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
os.makedirs(APP_DIR, exist_ok=True)

# 同期中・編集中に現れる一時ファイルは提出物として扱わない
IGNORED_PREFIXES = ("~$", ".~", ".")
IGNORED_SUFFIXES = (".tmp", ".partial", ".crdownload", ".lock", "~")


def _state_path(deadline_id: str) -> str:
    return os.path.join(APP_DIR, f"{deadline_id}_submissions.json")


def load_state(deadline_id: str) -> dict:
    """前回スキャン時のファイル一覧を返す。無ければ None。"""
    try:
        with open(_state_path(deadline_id), "r", encoding="utf-8") as f:
            state = json.load(f)
        if isinstance(state, dict) and isinstance(state.get("files"), dict):
            return state
    except Exception:
        pass
    return None


def save_state(deadline_id: str, files: dict):
    try:
        with open(_state_path(deadline_id), "w", encoding="utf-8") as f:
            json.dump(
                {"scanned_at": datetime.now().isoformat(timespec="seconds"), "files": files},
                f, ensure_ascii=False, indent=2
            )
    except Exception:
        pass


def _is_ignored(name: str) -> bool:
    lowered = name.lower()
    if lowered.startswith(IGNORED_PREFIXES):
        return True
    return lowered.endswith(IGNORED_SUFFIXES)


def parse_extensions(raw) -> list:
    """".xlsx, docx" のような入力を ['.xlsx', '.docx'] に正規化する。空なら制限なし。"""
    if not raw:
        return []
    if isinstance(raw, list):
        items = raw
    else:
        items = str(raw).replace("、", ",").split(",")
    exts = []
    for item in items:
        ext = item.strip().lower()
        if not ext:
            continue
        if not ext.startswith("."):
            ext = "." + ext
        exts.append(ext)
    return exts


def scan_folder(folder: str, recursive: bool = True, extensions=None) -> dict:
    """
    フォルダ内のファイルを {相対パス: {"mtime": float, "size": int}} 形式で返す。
    extensions を指定した場合はその拡張子のみ対象にする。
    """
    exts = parse_extensions(extensions)
    files = {}

    if recursive:
        walker = os.walk(folder)
    else:
        walker = [(folder, [], os.listdir(folder))]

    for root, _dirs, filenames in walker:
        for name in filenames:
            if _is_ignored(name):
                continue
            if exts and os.path.splitext(name)[1].lower() not in exts:
                continue
            full = os.path.join(root, name)
            try:
                stat = os.stat(full)
            except OSError:
                continue
            rel = os.path.relpath(full, folder).replace("\\", "/")
            files[rel] = {"mtime": stat.st_mtime, "size": stat.st_size}

    return files


def detect_changes(previous: dict, current: dict):
    """前回と今回のスキャン結果を比較し、(新規, 更新) のファイル名リストを返す。"""
    new_files, updated_files = [], []
    for rel, info in current.items():
        before = previous.get(rel)
        if before is None:
            new_files.append(rel)
        elif info["mtime"] > before.get("mtime", 0) or info["size"] != before.get("size"):
            updated_files.append(rel)
    return sorted(new_files), sorted(updated_files)


def build_message(folder: str, new_files: list, updated_files: list, mention: str = "") -> str:
    lines = []
    if mention:
        lines.append(mention)
    lines.append("📤 **提出フォルダに動きがありました**")
    lines.append(f"📁 {os.path.basename(os.path.normpath(folder)) or folder}")

    for rel in new_files:
        lines.append(f"・🆕 {rel}")
    for rel in updated_files:
        lines.append(f"・♻️ {rel}（更新）")

    body = "\n".join(lines)
    if len(body) > 1800:
        body = body[:1750] + "\n…（以下省略）"
    return body


def check_submissions(config: dict, log=print) -> dict:
    """
    設定に従って提出フォルダを確認し、新規/更新ファイルがあれば通知する。

    初回（状態ファイルが無い場合）は既存ファイルを記録するだけで通知しない。
    そうしないと、有効化した直後に既存ファイル全件が通知されてしまうため。

    戻り値: {"checked": bool, "new": [...], "updated": [...], "notified": bool}
    """
    result = {"checked": False, "new": [], "updated": [], "notified": False}

    if not config.get("submission_watch_enabled"):
        return result

    folder = config.get("submission_folder", "")
    if not folder or not os.path.isdir(folder):
        log(f"提出フォルダが見つかりません: {folder}")
        return result

    deadline_id = config.get("deadline_id", "default")
    current = scan_folder(
        folder,
        recursive=config.get("submission_recursive", True),
        extensions=config.get("submission_extensions", ""),
    )
    result["checked"] = True

    previous_state = load_state(deadline_id)
    if previous_state is None:
        save_state(deadline_id, current)
        log(f"提出フォルダの初回スキャンを記録しました（{len(current)}件、通知はしません）")
        return result

    new_files, updated_files = detect_changes(previous_state["files"], current)
    result["new"], result["updated"] = new_files, updated_files
    save_state(deadline_id, current)

    if not new_files and not updated_files:
        log("提出フォルダに変化はありません")
        return result

    webhook = config.get("submission_webhook_url", "") or config.get("webhook_url", "")
    if not webhook:
        log("提出通知先の Webhook URL がありません")
        return result

    message = build_message(folder, new_files, updated_files)
    try:
        response = send_webhook_text(webhook, message)
        status = getattr(response, "status_code", "N/A")
        log(f"提出通知を送信しました（新規{len(new_files)}件 / 更新{len(updated_files)}件）status={status}")
        result["notified"] = True
    except Exception as e:
        log(f"提出通知の送信に失敗しました: {e!r}")

    return result
