#update_checker.py
#
# GitHub の Releases を確認し、配布済みの締切教官より新しい版が
# 公開されていれば、起動時に知らせるためのモジュール。
#
# 配布側の運用:
#   1. 新しい SimekiriKyokan_Setup_vX.Y.exe をビルドする
#   2. GitHub の Releases でタグ vX.Y として公開し、exe を添付する
#   3. 利用者が次に起動したとき、このモジュールが新バージョンを検知して告知する

import re

import requests

# このビルドのバージョン。リリースのタグ名（vX.Y）と対応させる。
APP_VERSION = "2.1"

REPO = "RafuGH/SimekiriKyokan"
RELEASES_API = f"https://api.github.com/repos/{REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{REPO}/releases/latest"


def parse_version(text) -> tuple:
    """'v2.1.3' や '2.1' を (2, 1, 3) / (2, 1) のようなタプルにする。"""
    if not text:
        return ()
    numbers = re.findall(r"\d+", str(text))
    return tuple(int(n) for n in numbers)


def is_newer(latest, current) -> bool:
    """latest が current より新しいバージョンかどうか。"""
    lv, cv = parse_version(latest), parse_version(current)
    if not lv:
        return False
    # 桁数を揃えて比較する（2.1 と 2.1.1 など）
    length = max(len(lv), len(cv))
    lv += (0,) * (length - len(lv))
    cv += (0,) * (length - len(cv))
    return lv > cv


def fetch_latest_release(timeout: float = 6.0):
    """
    最新リリース情報を取得する。取得できない場合は None を返す
    （ネットワーク不通・レート制限・未公開などは想定内なので例外にしない）。
    """
    try:
        response = requests.get(
            RELEASES_API,
            timeout=timeout,
            headers={"Accept": "application/vnd.github+json"},
        )
        if response.status_code != 200:
            return None
        data = response.json()
    except Exception:
        return None

    if not isinstance(data, dict) or data.get("draft"):
        return None

    tag = data.get("tag_name") or ""
    if not tag:
        return None

    return {
        "version": tag,
        "name": data.get("name") or tag,
        "url": data.get("html_url") or RELEASES_PAGE,
        "body": (data.get("body") or "").strip(),
        "published_at": data.get("published_at") or "",
    }


def check_for_update(current_version: str = APP_VERSION):
    """
    新しいバージョンが公開されていればリリース情報を返す。
    最新版のまま・取得失敗の場合は None。
    """
    release = fetch_latest_release()
    if not release:
        return None
    if not is_newer(release["version"], current_version):
        return None
    return release


def format_message(release: dict, current_version: str = APP_VERSION) -> str:
    """告知ダイアログ用の本文を組み立てる。"""
    lines = [
        f"新しいバージョン {release['version']} が公開されています。",
        f"（お使いのバージョン: v{current_version}）",
    ]
    body = release.get("body", "")
    if body:
        summary = body if len(body) <= 400 else body[:400] + "…"
        lines += ["", "── 更新内容 ──", summary]
    lines += ["", "ダウンロードページを開きますか？"]
    return "\n".join(lines)
