#app_settings.py
#
# アプリ全体の設定（テーマ、アップデート確認など）の保存・読み込み。
# 締切ごとの設定(deadline_id.json)とは別に、アプリ共通の設定として扱う。

import json
import os

APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
os.makedirs(APP_DIR, exist_ok=True)

SETTINGS_PATH = os.path.join(APP_DIR, "app_settings.json")

DEFAULTS = {
    "theme": "light",             # "light" または "dark"
    "check_update_on_launch": True,
    "last_notified_version": "",   # 同じ新バージョンを毎回告知しないための記録
}


def load_settings() -> dict:
    """設定を読み込む。ファイルが無い/壊れている場合は既定値を返す。"""
    settings = dict(DEFAULTS)
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            stored = json.load(f)
        if isinstance(stored, dict):
            settings.update({k: v for k, v in stored.items() if k in DEFAULTS})
    except Exception:
        pass
    return settings


def save_settings(settings: dict):
    """設定を保存する。保存に失敗しても例外は投げない。"""
    try:
        current = load_settings()
        current.update({k: v for k, v in settings.items() if k in DEFAULTS})
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(current, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def get_theme() -> str:
    return load_settings().get("theme", DEFAULTS["theme"])


def set_theme(theme: str):
    save_settings({"theme": "dark" if theme == "dark" else "light"})


def is_dark() -> bool:
    return get_theme() == "dark"
