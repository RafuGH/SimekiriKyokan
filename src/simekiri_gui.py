# -*- coding: utf-8 -*-
# simekiri_gui.py  ── PyQt6 / Fluent-inspired design / Google OAuth2 / Help system

import ctypes
import sys, os, json, shutil, traceback
from datetime import datetime
import re, uuid

import simekiri_notify
import google_auth_helper

from PyQt6.QtWidgets import (
    QApplication, QWidget, QDialog, QScrollArea,
    QVBoxLayout, QHBoxLayout, QGridLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QSpinBox,
    QCheckBox, QDateEdit, QTimeEdit, QFileDialog,
    QMessageBox, QTableWidget, QTableWidgetItem,
    QSizePolicy, QFrame, QTextBrowser, QToolButton,
    QGraphicsDropShadowEffect
)
from PyQt6.QtCore import (
    QTime, QDate, Qt, QTimer, pyqtSlot,
    QSize, QPropertyAnimation, QEasingCurve
)
from PyQt6.QtGui import (
    QColor, QPalette, QFont, QFontDatabase,
    QIcon, QPixmap, QPainter, QBrush, QPen,
    QLinearGradient
)
from functools import partial


# ===================================================
# カラートークン（Fluent / Windowsライク）
# ===================================================

class Colors:
    # ライト
    L_BG        = "#f3f3f3"   # ウィンドウ背景
    L_SURFACE   = "#ffffff"   # カード・入力背景
    L_SURFACE2  = "#f9f9f9"   # サブ背景
    L_BORDER    = "#e0e0e0"   # ボーダー
    L_TEXT      = "#1b1b1b"   # メインテキスト
    L_TEXT2     = "#616161"   # サブテキスト
    L_ACCENT    = "#0078d4"   # Fluent アクセントブルー
    L_ACCENT_H  = "#106ebe"   # ホバー
    L_ACCENT_P  = "#005a9e"   # プレス
    L_DANGER    = "#c50f1f"   # 削除・警告
    L_SUCCESS   = "#107c10"   # 成功
    L_HEADER    = "#fafafa"   # タイトルバー背景

    # ダーク
    D_BG        = "#202020"
    D_SURFACE   = "#2c2c2c"
    D_SURFACE2  = "#383838"
    D_BORDER    = "#404040"
    D_TEXT      = "#ffffff"
    D_TEXT2     = "#9d9d9d"
    D_ACCENT    = "#60cdff"
    D_ACCENT_H  = "#4ec9f0"
    D_ACCENT_P  = "#3ab8e0"
    D_DANGER    = "#ff99a4"
    D_SUCCESS   = "#6ccb5f"
    D_HEADER    = "#1c1c1c"


def make_stylesheet(dark: bool) -> str:
    c = Colors
    if dark:
        bg, surf, surf2, brd = c.D_BG, c.D_SURFACE, c.D_SURFACE2, c.D_BORDER
        txt, txt2            = c.D_TEXT, c.D_TEXT2
        acc, acc_h, acc_p    = c.D_ACCENT, c.D_ACCENT_H, c.D_ACCENT_P
        danger, success      = c.D_DANGER, c.D_SUCCESS
        hdr                  = c.D_HEADER
    else:
        bg, surf, surf2, brd = c.L_BG, c.L_SURFACE, c.L_SURFACE2, c.L_BORDER
        txt, txt2            = c.L_TEXT, c.L_TEXT2
        acc, acc_h, acc_p    = c.L_ACCENT, c.L_ACCENT_H, c.L_ACCENT_P
        danger, success      = c.L_DANGER, c.L_SUCCESS
        hdr                  = c.L_HEADER

    return f"""
/* ── ベース ── */
QWidget {{
    font-family: "Segoe UI", "Meiryo UI", "Yu Gothic UI", sans-serif;
    font-size: 13px;
    color: {txt};
    background-color: {bg};
}}
QScrollArea, QScrollArea > QWidget > QWidget {{ background: {bg}; border: none; }}

/* ── タイトルバー相当 ── */
QWidget#titlebar {{
    background-color: {hdr};
    border-bottom: 1px solid {brd};
}}

/* ── カードフレーム ── */
QFrame#card {{
    background-color: {surf};
    border: 1px solid {brd};
    border-radius: 8px;
}}

/* ── セクションヘッダー ── */
QLabel#section_hdr {{
    font-size: 11px;
    font-weight: 600;
    color: {txt2};
    letter-spacing: 0.8px;
    padding: 14px 0 2px 0;
}}

/* ── フィールドラベル ── */
QLabel#field_lbl {{
    font-size: 12px;
    color: {txt2};
    padding: 4px 0 1px 0;
}}

/* ── 説明ラベル ── */
QLabel#desc_lbl {{
    font-size: 11px;
    color: {txt2};
    padding: 0;
}}

/* ── 汎用ラベル ── */
QLabel {{ color: {txt}; background: transparent; }}

/* ── 入力フィールド共通 ── */
QLineEdit, QSpinBox, QComboBox, QTimeEdit, QDateEdit {{
    background: {surf};
    border: 1px solid {brd};
    border-radius: 4px;
    padding: 5px 10px;
    color: {txt};
    min-height: 30px;
    selection-background-color: {acc};
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus,
QTimeEdit:focus, QDateEdit:focus {{
    border: 1.5px solid {acc};
}}
QLineEdit::placeholder {{ color: {txt2}; }}
QComboBox::drop-down {{ border: none; width: 20px; }}
QComboBox QAbstractItemView {{
    background: {surf};
    border: 1px solid {brd};
    selection-background-color: {acc};
    color: {txt};
}}

/* ── プライマリボタン ── */
QPushButton#btn_primary {{
    background-color: {acc};
    color: {"#000000" if dark else "#ffffff"};
    border: none;
    border-radius: 4px;
    padding: 7px 20px;
    font-weight: 600;
    min-height: 32px;
}}
QPushButton#btn_primary:hover  {{ background-color: {acc_h}; }}
QPushButton#btn_primary:pressed {{ background-color: {acc_p}; }}
QPushButton#btn_primary:disabled {{ background-color: {brd}; color: {txt2}; }}

/* ── セカンダリボタン ── */
QPushButton {{
    background-color: {surf};
    color: {txt};
    border: 1px solid {brd};
    border-radius: 4px;
    padding: 6px 16px;
    min-height: 32px;
}}
QPushButton:hover  {{ background-color: {surf2}; border-color: {txt2}; }}
QPushButton:pressed {{ background-color: {brd}; }}
QPushButton:disabled {{ color: {txt2}; }}

/* ── 危険ボタン ── */
QPushButton#btn_danger {{
    color: {danger};
    border-color: {danger};
    background: transparent;
}}
QPushButton#btn_danger:hover {{ background-color: {"#3a1010" if dark else "#fff0f0"}; }}

/* ── ヘルプボタン ── */
QToolButton#btn_help {{
    background: transparent;
    border: 1px solid {brd};
    border-radius: 10px;
    color: {txt2};
    font-size: 11px;
    font-weight: 600;
    min-width: 20px;
    max-width: 20px;
    min-height: 20px;
    max-height: 20px;
    padding: 0;
}}
QToolButton#btn_help:hover {{ border-color: {acc}; color: {acc}; background: {surf2}; }}

/* ── チェックボックス ── */
QCheckBox {{ spacing: 8px; color: {txt}; }}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border: 1px solid {brd};
    border-radius: 3px;
    background: {surf};
}}
QCheckBox::indicator:checked {{
    background: {acc};
    border-color: {acc};
}}

/* ── テーブル ── */
QTableWidget {{
    background: {surf};
    border: 1px solid {brd};
    border-radius: 6px;
    gridline-color: {brd};
    outline: none;
}}
QTableWidget::item {{ padding: 6px 10px; }}
QTableWidget::item:selected {{
    background: {"#004f8c" if dark else "#cce4f7"};
    color: {txt};
}}
QTableWidget::item:alternate {{ background: {surf2}; }}
QHeaderView::section {{
    background: {surf2};
    border: none;
    border-bottom: 1px solid {brd};
    padding: 6px 10px;
    font-weight: 600;
    font-size: 11px;
    color: {txt2};
}}

/* ── スクロールバー ── */
QScrollBar:vertical {{
    width: 8px; background: transparent; margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {brd}; border-radius: 4px; min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {txt2}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}

/* ── テキストブラウザ（ヘルプウィンドウ） ── */
QTextBrowser {{
    background: {surf};
    border: none;
    color: {txt};
    font-size: 13px;
    line-height: 1.6;
}}

/* ── セパレータ ── */
QFrame[frameShape="4"] {{ color: {brd}; }}
"""


# ===================================================
# ヘルプテキスト辞書
# ===================================================

HELP_TEXTS = {
    "title": (
        "締切名",
        "このプロジェクトや締切の名前を入力します。\n\n"
        "例：「夏コミ新作ゲーム」「卒業論文」\n\n"
        "タスク管理画面での識別に使われます。"
    ),
    "category": (
        "カテゴリ",
        "プロジェクトのカテゴリを選択します。\n\n"
        "• game ── ゲーム制作\n"
        "• report ── レポート・論文\n"
        "• school ── 学校課題\n"
        "• work ── 仕事・業務\n"
        "• personal ── 個人プロジェクト"
    ),
    "datasource": (
        "データソース",
        "タスクデータの読み込み元を選択します。\n\n"
        "【Excel ファイル】\n"
        "ローカルの .xlsx ファイルを使用します。\n"
        "「テンプレート Excel を生成」でひな型を作れます。\n\n"
        "【Google スプレッドシート】\n"
        "Google Drive 上のスプレッドシートを使用します。\n"
        "「Google で認可する」ボタンで事前に認証が必要です。"
    ),
    "google_auth": (
        "Google 認証について",
        "Google Sheets を使うには一度だけ認証が必要です。\n\n"
        "【手順】\n"
        "1. Google Cloud Console でプロジェクトを作成\n"
        "2. Google Sheets API を有効化\n"
        "3. OAuth 2.0 クライアントIDを作成\n"
        "   （リダイレクト URI: http://localhost:8080/callback）\n"
        "4. credentials.json をダウンロードして以下に配置：\n"
        "   %LOCALAPPDATA%\\SimekiriKyokan\\credentials.json\n\n"
        "【403エラーが出る場合】\n"
        "Google Cloud Console → OAuth同意画面 →\n"
        "「テストユーザー」にあなたの Gmail を追加してください。"
    ),
    "webhook": (
        "Webhook URL",
        "通知を送信する先の Webhook URL を入力します。\n\n"
        "URLから送信先が自動判定されます：\n"
        "• Discord ── discord.com/api/webhooks/...\n"
        "• Slack ── hooks.slack.com/...\n"
        "• Teams ── outlook.office.com/webhook/...\n"
        "• Chatwork ── chatwork.com/...\n"
        "• Google Chat ── chat.googleapis.com/...\n\n"
        "画像送信対応：Discord・Teams・Google Chat"
    ),
    "days_before": (
        "締切前通知日数",
        "締切の何日前から通知を送り始めるかを設定します。\n\n"
        "• 0 ── 当日と過ぎた分のみ\n"
        "• 3 ── 3日前〜当日（推奨）\n"
        "• 7 ── 1週間前から\n\n"
        "締切を過ぎたタスクも常に通知されます。"
    ),
    "mention": (
        "メンション設定",
        "通知メッセージに担当者へのメンションを付けます。\n\n"
        "「担当名」── Excelの「担当」列に入力されている名前\n"
        "「ユーザーID」── 各プラットフォームのID\n\n"
        "Discord の場合：ユーザーIDまたはユーザー名\n"
        "Slack の場合：UXXXXXXXX 形式のメンバーID\n"
        "Teams の場合：表示名\n"
        "Chatwork の場合：数字のユーザーID"
    ),
    "reviewer": (
        "確認待ち通知",
        "Excelで進捗が「確認待ち」になったタスクを、\n"
        "担当者ではなくレビュアーに通知します。\n\n"
        "「確認待ち」タスクは通常の締切通知からは除外され、\n"
        "レビュアー専用の通知として別途送信されます。\n\n"
        "Webhook URL を空欄にすると、メインと同じURLを使用します。"
    ),
    "auto_notify": (
        "自動連絡",
        "Windowsのタスクスケジューラを使って、\n"
        "指定した時刻に自動で通知を送信します。\n\n"
        "【注意】登録には管理者権限が必要です。\n\n"
        "• 送信時刻 ── 毎日この時刻に通知\n"
        "• 送信頻度 ── N日おきに送信（1=毎日）\n"
        "• 制作期間 ── この期間中のみ自動送信"
    ),
}


# ===================================================
# ユーティリティ関数
# ===================================================

def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

ShellExecuteW  = ctypes.windll.shell32.ShellExecuteW
APP_DIR        = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
os.makedirs(APP_DIR, exist_ok=True)
ERROR_LOG      = os.path.join(APP_DIR, "gui_error_log.txt")
TASK_BASE_NAME = "SimekiriKyokan"
ADMIN_FLAG     = "--admin-register"

def get_config_path(did):      return os.path.join(APP_DIR, f"{did}.json")
def get_task_config_path(did): return os.path.join(APP_DIR, f"{did}.json")
def get_task_name(did):        return f"{TASK_BASE_NAME}_{did}"

def generate_deadline_id(cat, end, title):
    safe = re.sub(r'[^a-zA-Z0-9ぁ-んァ-ン一-龯]', '', title)[:10]
    return f"{cat}_{safe}_{uuid.uuid4().hex[:6]}"

def task_exists(did):
    try:
        import win32com.client
        s = win32com.client.Dispatch("Schedule.Service"); s.Connect()
        s.GetFolder("\\").GetTask(get_task_name(did)); return True
    except: return False

def get_simekiri_tasks():
    import win32com.client
    s = win32com.client.Dispatch("Schedule.Service"); s.Connect()
    return [
        {"name": t.Name, "state": t.State, "enabled": t.Enabled,
         "next_run": str(t.NextRunTime), "last_run": str(t.LastRunTime),
         "last_result": t.LastTaskResult}
        for t in s.GetFolder("\\").GetTasks(0)
        if t.Name.startswith(TASK_BASE_NAME)
    ]

def register_task_admin(cfg):
    import win32com.client
    tn = get_task_name(cfg["deadline_id"])
    s = win32com.client.Dispatch("Schedule.Service"); s.Connect()
    root = s.GetFolder("\\")
    try: root.DeleteTask(tn, 0)
    except: pass
    td = s.NewTask(0)
    h, m = map(int, cfg["notify_time"].split(":"))
    st = datetime.strptime(cfg["start_date"], "%Y-%m-%d").replace(hour=h, minute=m)
    en = datetime.strptime(cfg["end_date"],   "%Y-%m-%d").replace(hour=23, minute=59, second=59)
    tr = td.Triggers.Create(2)
    tr.StartBoundary = st.strftime("%Y-%m-%dT%H:%M:%S")
    tr.EndBoundary   = en.strftime("%Y-%m-%dT%H:%M:%S")
    tr.DaysInterval  = max(1, cfg["notify_interval_days"]); tr.Enabled = True
    ac = td.Actions.Create(0)
    cp = get_config_path(cfg["deadline_id"])
    if getattr(sys, "frozen", False):
        ac.Path = sys.executable; ac.Arguments = f'--notify "{cp}"'
        ac.WorkingDirectory = os.path.dirname(sys.executable)
    else:
        ac.Path = sys.executable
        ac.Arguments = f'"{os.path.abspath(__file__)}" --notify "{cp}"'
        ac.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
    td.Principal.LogonType = 3; td.Principal.RunLevel = 0
    td.Settings.Enabled = True; td.Settings.StartWhenAvailable = True
    td.Settings.ExecutionTimeLimit = "PT0S"
    root.RegisterTaskDefinition(tn, td, 6, None, None, 3)
    cfg["task_registered"] = True


# ===================================================
# 共通ウィジェット：ヘルプボタン
# ===================================================

class HelpButton(QToolButton):
    """「?」マークの小さなヘルプボタン"""
    def __init__(self, key: str, parent=None):
        super().__init__(parent)
        self.help_key = key
        self.setText("?")
        self.setObjectName("btn_help")
        self.setFixedSize(20, 20)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clicked.connect(self._show_help)

    def _show_help(self):
        title, body = HELP_TEXTS.get(self.help_key, ("ヘルプ", "説明がありません。"))
        dlg = HelpDialog(title, body, self.window())
        dlg.exec()


class HelpDialog(QDialog):
    """ヘルプ表示ダイアログ"""
    def __init__(self, title: str, body: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"ヘルプ – {title}")
        self.setFixedWidth(420)
        self.setWindowFlags(
            Qt.WindowType.Dialog |
            Qt.WindowType.WindowCloseButtonHint
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        hdr = QLabel(title)
        hdr.setStyleSheet("font-size:15px; font-weight:600;")
        layout.addWidget(hdr)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)

        browser = QTextBrowser()
        browser.setPlainText(body)
        browser.setMinimumHeight(120)
        browser.setMaximumHeight(320)
        layout.addWidget(browser)

        close_btn = QPushButton("閉じる")
        close_btn.setObjectName("btn_primary")
        close_btn.setFixedWidth(80)
        close_btn.clicked.connect(self.accept)
        row = QHBoxLayout(); row.addStretch(); row.addWidget(close_btn)
        layout.addLayout(row)


# ===================================================
# フィールド行ヘルパー（ラベル + ヘルプボタン）
# ===================================================

def field_row(label_text: str, help_key: str = None) -> QHBoxLayout:
    """「ラベル ─ ヘルプボタン」の横並びレイアウトを返す"""
    row = QHBoxLayout(); row.setContentsMargins(0, 6, 0, 2)
    lbl = QLabel(label_text); lbl.setObjectName("field_lbl")
    row.addWidget(lbl)
    if help_key:
        row.addWidget(HelpButton(help_key))
    row.addStretch()
    return row


def section_header(text: str) -> QLabel:
    lbl = QLabel(text.upper()); lbl.setObjectName("section_hdr")
    return lbl


# ===================================================
# Google アカウント表示ウィジェット
# ===================================================

class AccountBadge(QWidget):
    """ヘッダー右端に表示するアカウントバッジ"""
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 4, 0)
        layout.setSpacing(6)

        self.icon_lbl = QLabel()
        self.icon_lbl.setFixedSize(22, 22)
        self.name_lbl = QLabel()
        self.name_lbl.setObjectName("desc_lbl")

        layout.addWidget(self.icon_lbl)
        layout.addWidget(self.name_lbl)
        self.refresh()

    def refresh(self):
        if google_auth_helper.has_token():
            # トークンからメールを取得
            try:
                creds = google_auth_helper.get_creds()
                import json as _json
                token_path = google_auth_helper.TOKEN_PATH
                with open(token_path, "r") as f:
                    data = _json.load(f)
                email = data.get("client_id", "").split("-")[0] if "client_id" in data else ""
                # id_token からメールを取得試み
                id_tok = data.get("id_token", "")
                if id_tok:
                    import base64
                    payload = id_tok.split(".")[1]
                    payload += "=" * (-len(payload) % 4)
                    info = _json.loads(base64.b64decode(payload))
                    email = info.get("email", email)
                display = email if email else "認証済み"
            except Exception:
                display = "認証済み"

            self._draw_icon(True)
            self.name_lbl.setText(display)
            self.setToolTip(f"Google アカウント: {display}")
        else:
            self._draw_icon(False)
            self.name_lbl.setText("未ログイン")
            self.setToolTip("Google Sheets を使う場合は認可が必要です")

    def _draw_icon(self, logged_in: bool):
        pix = QPixmap(22, 22)
        pix.fill(Qt.GlobalColor.transparent)
        p = QPainter(pix)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        color = QColor("#107c10") if logged_in else QColor("#797979")
        p.setBrush(QBrush(color))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(0, 0, 22, 22)
        p.setPen(QPen(QColor("white")))
        p.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        p.drawText(pix.rect(), Qt.AlignmentFlag.AlignCenter, "G" if logged_in else "—")
        p.end()
        self.icon_lbl.setPixmap(pix)


# ===================================================
# メンション行入力
# ===================================================

class RowInput(QWidget):
    def __init__(self, short_ph, long_ph, parent_layout, deletable=True):
        super().__init__()
        self.parent_layout = parent_layout
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2); layout.setSpacing(6)
        self.short = QLineEdit(); self.short.setPlaceholderText(short_ph)
        self.long  = QLineEdit(); self.long.setPlaceholderText(long_ph)
        self.add_btn = QPushButton("＋"); self.add_btn.setFixedWidth(34)
        self.del_btn = QPushButton("－"); self.del_btn.setFixedWidth(34)
        self.add_btn.clicked.connect(self.add)
        self.del_btn.clicked.connect(self.delete)
        if not deletable: self.del_btn.setEnabled(False)
        for w in (self.short, self.long, self.add_btn, self.del_btn):
            layout.addWidget(w)

    def update_delete_state(self):
        self.del_btn.setEnabled(self.parent_layout.count() > 1)

    def add(self):
        r = RowInput(self.short.placeholderText(), self.long.placeholderText(), self.parent_layout)
        self.parent_layout.addWidget(r); self.update_all()

    def delete(self):
        if self.parent_layout.count() <= 1: return
        self.parent_layout.removeWidget(self); self.deleteLater(); self.update_all()

    def update_all(self):
        for i in range(self.parent_layout.count()):
            w = self.parent_layout.itemAt(i).widget()
            if isinstance(w, RowInput): w.update_delete_state()

    def get(self): return self.short.text(), self.long.text()


# ===================================================
# Google 認可ミックスイン
# ===================================================

class GoogleAuthMixin:
    def _authorize_google(self):
        self.auth_btn.setEnabled(False)
        self.auth_status_label.setText("🔄 ブラウザで認可中…")
        def _cb(ok, msg):
            QTimer.singleShot(0, lambda: self._on_auth_done(ok, msg))
        google_auth_helper.authorize_google_sheets(_cb)

    def _on_auth_done(self, ok, msg):
        if ok: QMessageBox.information(self, "Google 認証", msg)
        else:   QMessageBox.warning(self,   "Google 認証", msg)
        self._refresh_auth_label()
        # アカウントバッジを更新
        if hasattr(self, "account_badge"):
            self.account_badge.refresh()

    def _refresh_auth_label(self):
        ok = google_auth_helper.has_token()
        self.auth_status_label.setText("✅ 認可済み" if ok else "❌ 未認可")
        self.auth_btn.setEnabled(not ok)


# ===================================================
# メインGUI
# ===================================================

class NotifierApp(GoogleAuthMixin, QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("締切教官")
        self.setMinimumWidth(560)
        self.resize(580, 880)

        dark = self.palette().color(QPalette.ColorRole.Window).lightness() < 128
        self.setStyleSheet(make_stylesheet(dark))

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---- タイトルバー ----
        titlebar = QWidget(); titlebar.setObjectName("titlebar")
        tb_layout = QHBoxLayout(titlebar)
        tb_layout.setContentsMargins(16, 10, 12, 10)
        app_lbl = QLabel("締切教官")
        app_lbl.setStyleSheet("font-size:15px; font-weight:700; letter-spacing:0.5px;")
        self.account_badge = AccountBadge()
        tb_layout.addWidget(app_lbl)
        tb_layout.addStretch()
        tb_layout.addWidget(self.account_badge)
        root.addWidget(titlebar)

        # ---- スクロール本体 ----
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        inner  = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(20, 16, 20, 24)
        layout.setSpacing(2)
        scroll.setWidget(inner)
        root.addWidget(scroll)

        # ━━ 基本情報 ━━
        layout.addWidget(section_header("基本情報"))

        layout.addLayout(field_row("締切名", "title"))
        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText("例：夏コミ新作ゲーム")
        layout.addWidget(self.title_input)

        layout.addLayout(field_row("カテゴリ", "category"))
        self.category_combo = QComboBox()
        self.category_combo.addItems(["game", "report", "school", "work", "personal"])
        layout.addWidget(self.category_combo)

        # ━━ データソース ━━
        layout.addWidget(section_header("データソース"))
        layout.addLayout(field_row("読み込み元", "datasource"))
        self.source_combo = QComboBox()
        self.source_combo.addItems(["Excel ファイル", "Google スプレッドシート"])
        self.source_combo.currentIndexChanged.connect(self._on_source_changed)
        layout.addWidget(self.source_combo)

        # Excel 欄
        self.excel_group = QWidget()
        evl = QVBoxLayout(self.excel_group)
        evl.setContentsMargins(0, 4, 0, 0); evl.setSpacing(4)
        evl.addLayout(field_row("Excel ファイルのパス"))
        ehl = QHBoxLayout(); ehl.setSpacing(6)
        self.excel_input = QLineEdit()
        self.excel_input.setPlaceholderText("Tasks.xlsx を選択…")
        btn_e = QPushButton("参照"); btn_e.setFixedWidth(64)
        btn_e.clicked.connect(self.browse_excel)
        ehl.addWidget(self.excel_input); ehl.addWidget(btn_e)
        evl.addLayout(ehl)
        self.gen_excel_btn = QPushButton("テンプレート Excel を生成")
        self.gen_excel_btn.clicked.connect(self.generate_excel)
        evl.addWidget(self.gen_excel_btn)
        layout.addWidget(self.excel_group)

        # Sheets 欄
        self.sheets_group = QWidget()
        svl = QVBoxLayout(self.sheets_group)
        svl.setContentsMargins(0, 4, 0, 0); svl.setSpacing(4)
        svl.addLayout(field_row("スプレッドシート URL", "google_auth"))
        self.sheets_url_input = QLineEdit()
        self.sheets_url_input.setPlaceholderText("https://docs.google.com/spreadsheets/d/...")
        svl.addWidget(self.sheets_url_input)
        auth_hl = QHBoxLayout(); auth_hl.setSpacing(10)
        self.auth_btn = QPushButton("🔐  Google で認可する")
        self.auth_btn.setObjectName("btn_primary")
        self.auth_btn.clicked.connect(self._authorize_google)
        self.auth_status_label = QLabel()
        auth_hl.addWidget(self.auth_btn)
        auth_hl.addWidget(self.auth_status_label)
        auth_hl.addStretch()
        svl.addLayout(auth_hl)
        layout.addWidget(self.sheets_group)
        self.sheets_group.setVisible(False)
        self._refresh_auth_label()

        # ━━ 通知先 ━━
        layout.addWidget(section_header("通知先"))
        layout.addLayout(field_row("Webhook URL", "webhook"))
        self.webhook_input = QLineEdit()
        self.webhook_input.setPlaceholderText("https://discord.com/api/webhooks/...")
        layout.addWidget(self.webhook_input)

        layout.addLayout(field_row("締切何日前に通知", "days_before"))
        self.days_spin = QSpinBox(); self.days_spin.setRange(0, 60)
        self.days_spin.setFixedWidth(100)
        layout.addWidget(self.days_spin)

        # メンション
        layout.addLayout(field_row("担当者メンション（任意）", "mention"))
        self.mention_checkbox = QCheckBox("メンションを有効にする")
        layout.addWidget(self.mention_checkbox)
        mb = QWidget()
        ml = QVBoxLayout(mb); ml.setContentsMargins(20,0,0,0); ml.setSpacing(2)
        self.mention_layout = ml
        self.mention_layout.addWidget(RowInput("担当名", "ユーザーID", self.mention_layout, False))
        layout.addWidget(mb)

        # ━━ 確認待ち通知 ━━
        layout.addWidget(section_header("確認待ち通知"))
        layout.addLayout(field_row("レビュアーへの通知（任意）", "reviewer"))
        self.reviewer_checkbox = QCheckBox("確認待ちタスクを別の担当者に通知する")
        layout.addWidget(self.reviewer_checkbox)
        self.reviewer_checkbox.stateChanged.connect(self._on_reviewer_toggled)

        self.reviewer_group = QWidget()
        rvl = QVBoxLayout(self.reviewer_group)
        rvl.setContentsMargins(20, 4, 0, 0); rvl.setSpacing(4)
        rvl.addLayout(field_row("レビュアー Webhook URL（省略可）"))
        self.reviewer_webhook_input = QLineEdit()
        self.reviewer_webhook_input.setPlaceholderText("省略すると上の URL を使用")
        rvl.addWidget(self.reviewer_webhook_input)
        rvl.addLayout(field_row("レビュアーのメンション"))
        self.reviewer_mention_layout = QVBoxLayout()
        self.reviewer_mention_layout.setSpacing(2)
        self.reviewer_mention_layout.addWidget(
            RowInput("担当名", "レビュアーID", self.reviewer_mention_layout, False)
        )
        rvm = QWidget(); rvm.setLayout(self.reviewer_mention_layout)
        rvl.addWidget(rvm)
        layout.addWidget(self.reviewer_group)
        self.reviewer_group.setVisible(False)

        # ━━ 自動連絡 ━━
        layout.addWidget(section_header("自動連絡"))
        layout.addLayout(field_row("タスクスケジューラ連携（任意）", "auto_notify"))
        self.auto_checkbox = QCheckBox("自動送信を有効にする")
        layout.addWidget(self.auto_checkbox)

        tl = QHBoxLayout(); tl.setSpacing(10)
        tl.addWidget(QLabel("送信時刻"))
        self.time_edit = QTimeEdit(QTime(9, 0))
        self.time_edit.setFixedWidth(100)
        tl.addWidget(self.time_edit)
        tl.addSpacing(20)
        tl.addWidget(QLabel("頻度（日おき）"))
        self.interval_spin = QSpinBox(); self.interval_spin.setRange(1, 30)
        self.interval_spin.setFixedWidth(80)
        tl.addWidget(self.interval_spin); tl.addStretch()
        layout.addLayout(tl)

        layout.addLayout(field_row("制作期間"))
        dl = QHBoxLayout(); dl.setSpacing(8)
        self.start_date = QDateEdit(QDate.currentDate())
        self.start_date.setCalendarPopup(True)
        self.end_date = QDateEdit(QDate.currentDate().addYears(1))
        self.end_date.setCalendarPopup(True)
        dl.addWidget(QLabel("開始")); dl.addWidget(self.start_date)
        dl.addSpacing(8)
        dl.addWidget(QLabel("終了")); dl.addWidget(self.end_date)
        dl.addStretch()
        layout.addLayout(dl)

        # ━━ アクションボタン ━━
        layout.addSpacing(16)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)
        layout.addSpacing(10)

        btn_row1 = QHBoxLayout(); btn_row1.setSpacing(8)
        self.save_btn = QPushButton("新規作成して保存")
        self.save_btn.setObjectName("btn_primary")
        self.save_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        btn_row1.addWidget(self.save_btn)
        layout.addLayout(btn_row1)

        btn_row2 = QHBoxLayout(); btn_row2.setSpacing(8)
        self.run_btn  = QPushButton("📨  通知テストを送信")
        self.list_btn = QPushButton("⚙  締切教官を管理")
        for b in (self.run_btn, self.list_btn):
            b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            btn_row2.addWidget(b)
        layout.addLayout(btn_row2)
        layout.addStretch()

        self.save_btn.clicked.connect(self.save_config)
        self.run_btn.clicked.connect(self.run_notify)
        self.list_btn.clicked.connect(self.open_task_list)
        self.config = {}

    # ---- ソース切り替え ----
    def _on_source_changed(self, i):
        self.excel_group.setVisible(i == 0)
        self.sheets_group.setVisible(i == 1)

    def _on_reviewer_toggled(self, state):
        self.reviewer_group.setVisible(state == Qt.CheckState.Checked.value)

    def browse_excel(self):
        p, _ = QFileDialog.getOpenFileName(self, "Excel を選択", "", "Excel (*.xlsx *.xls)")
        if p: self.excel_input.setText(p)

    def open_manual(self):
        base = os.path.dirname(sys.executable if getattr(sys,"frozen",False) else os.path.abspath(__file__))
        pdf  = os.path.join(base, "SimekiriKyokan_Manual.pdf")
        if os.path.exists(pdf): os.startfile(pdf)
        else: QMessageBox.warning(self, "エラー", "マニュアル PDF が見つかりません")

    def generate_excel(self):
        base = os.path.dirname(sys.executable if getattr(sys,"frozen",False) else os.path.abspath(__file__))
        tmpl = os.path.join(base, "Tasks.xlsx")
        if not os.path.exists(tmpl):
            QMessageBox.critical(self,"エラー",f"テンプレートが見つかりません:\n{tmpl}"); return
        dst, _ = QFileDialog.getSaveFileName(self, "保存先を選択", "Tasks.xlsx", "Excel (*.xlsx)")
        if not dst: return
        if os.path.exists(dst):
            if QMessageBox.question(self,"上書き確認",f"上書きしますか？\n{dst}",
                    QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
                return
        shutil.copyfile(tmpl, dst)
        QMessageBox.information(self,"完了",f"生成しました:\n{dst}")

    def _collect_mentions(self, lyt):
        res = []
        for i in range(lyt.count()):
            w = lyt.itemAt(i).widget()
            if isinstance(w, RowInput):
                s,l = w.get()
                if s or l: res.append({"name":s,"id":l})
        return res

    def save_config(self):
        title = self.title_input.text().strip()
        if not title:
            QMessageBox.warning(self,"入力エラー","締切名を入力してください"); return
        is_sheets = (self.source_combo.currentIndex() == 1)
        if is_sheets:
            if not self.sheets_url_input.text().strip():
                QMessageBox.warning(self,"入力エラー","スプレッドシート URL を入力してください"); return
            if not google_auth_helper.has_token():
                QMessageBox.warning(self,"入力エラー","Google 認証を先に完了してください"); return
        else:
            if not self.excel_input.text().strip():
                QMessageBox.warning(self,"入力エラー","Excel ファイルを指定してください"); return
        if not self.webhook_input.text().strip():
            QMessageBox.warning(self,"入力エラー","Webhook URL を入力してください"); return

        cat = self.category_combo.currentText()
        ed  = self.end_date.date().toString("yyyy-MM-dd")
        did = generate_deadline_id(cat, ed, title)

        cfg = {
            "deadline_id": did, "title": title, "category": cat,
            "data_source": "sheets" if is_sheets else "excel",
            "excel_path":  self.excel_input.text() if not is_sheets else "",
            "sheets_url":  self.sheets_url_input.text() if is_sheets else "",
            "webhook_url":          self.webhook_input.text(),
            "days_before_deadline": self.days_spin.value(),
            "mention_enabled":      self.mention_checkbox.isChecked(),
            "mentions":             self._collect_mentions(self.mention_layout),
            "reviewer_enabled":     self.reviewer_checkbox.isChecked(),
            "reviewer_webhook_url": self.reviewer_webhook_input.text(),
            "reviewer_mentions":    self._collect_mentions(self.reviewer_mention_layout),
            "auto_notify":          self.auto_checkbox.isChecked(),
            "notify_time":          self.time_edit.time().toString("HH:mm"),
            "notify_interval_days": self.interval_spin.value(),
            "start_date":           self.start_date.date().toString("yyyy-MM-dd"),
            "end_date":             ed,
        }
        cp = get_config_path(did)
        with open(cp,"w",encoding="utf-8") as f: json.dump(cfg,f,ensure_ascii=False,indent=2)
        self.config = cfg

        if cfg["auto_notify"]:
            try:
                import win32com.client
                if is_admin():
                    register_task_admin(cfg)
                    QMessageBox.information(self,"保存完了","設定を保存してタスクを登録しました")
                else:
                    if QMessageBox.question(self,"管理者権限",
                            "タスク登録には管理者権限が必要です。昇格しますか？",
                            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
                        exe = sys.executable
                        fl  = f'{ADMIN_FLAG} "{cp}"'
                        if not getattr(sys,"frozen",False): fl = f'"{os.path.abspath(__file__)}" ' + fl
                        ShellExecuteW(None,"runas",exe,fl,None,1)
                        QMessageBox.information(self,"保存完了","設定を保存しました（管理者認証待ち）")
                    else:
                        QMessageBox.information(self,"保存完了","設定を保存しました（タスク未登録）")
            except Exception as e:
                QMessageBox.warning(self,"エラー",f"タスク登録に失敗しました: {e}")
        else:
            QMessageBox.information(self,"保存完了","設定を保存しました")
        self._reset_form()

    def _reset_form(self):
        self.config = {}
        for w in (self.title_input, self.excel_input, self.sheets_url_input,
                  self.webhook_input, self.reviewer_webhook_input):
            w.clear()
        self.days_spin.setValue(3)
        self.mention_checkbox.setChecked(False)
        self.reviewer_checkbox.setChecked(False)
        self.auto_checkbox.setChecked(False)
        self.time_edit.setTime(QTime(9,0))
        self.interval_spin.setValue(1)
        self.source_combo.setCurrentIndex(0)
        for lyt in (self.mention_layout, self.reviewer_mention_layout):
            while lyt.count():
                w = lyt.takeAt(0).widget()
                if w: w.setParent(None)
        self.mention_layout.addWidget(RowInput("担当名","ユーザーID",self.mention_layout,False))
        self.reviewer_mention_layout.addWidget(RowInput("担当名","レビュアーID",self.reviewer_mention_layout,False))

    def run_notify(self):
        files = sorted([f for f in os.listdir(APP_DIR) if f.endswith(".json")],
                       key=lambda x: os.path.getmtime(os.path.join(APP_DIR,x)),reverse=True)
        if not files: QMessageBox.warning(self,"エラー","保存された設定がありません"); return
        simekiri_notify.run_notify(os.path.join(APP_DIR,files[0]),test_mode=True)

    def open_task_list(self):
        self.task_list_window = TaskManagerWindow(self)
        self.task_list_window.show()

    def update_task(self, cfg):
        tn = get_task_name(cfg["deadline_id"]); cp = get_config_path(cfg["deadline_id"])
        if is_admin():
            try:
                import win32com.client
                s = win32com.client.Dispatch("Schedule.Service"); s.Connect()
                try: s.GetFolder("\\").DeleteTask(tn,0)
                except: pass
                register_task_admin(cfg)
            except Exception as e: QMessageBox.warning(self,"タスク更新失敗",str(e))
            return
        fl = f'{ADMIN_FLAG} "{cp}"'
        if not getattr(sys,"frozen",False): fl = f'"{os.path.abspath(__file__)}" ' + fl
        ShellExecuteW(None,"runas",sys.executable,fl,None,1)


# ===================================================
# タスク管理ウィンドウ
# ===================================================

class TaskManagerWindow(QWidget):
    def __init__(self, main_app):
        super().__init__()
        self.main_app = main_app
        self.setWindowTitle("締切教官 – 管理")
        self.resize(900, 460)

        dark = self.palette().color(QPalette.ColorRole.Window).lightness() < 128
        self.setStyleSheet(make_stylesheet(dark))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20,16,20,20); layout.setSpacing(12)

        hdr_row = QHBoxLayout()
        hdr = QLabel("登録済みの教官一覧")
        hdr.setStyleSheet("font-size:15px; font-weight:700;")
        self.refresh_btn = QPushButton("🔄  更新")
        self.refresh_btn.setFixedWidth(90)
        self.refresh_btn.clicked.connect(self.load_tasks)
        hdr_row.addWidget(hdr); hdr_row.addStretch(); hdr_row.addWidget(self.refresh_btn)
        layout.addLayout(hdr_row)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["教官名","状態","次回実行","最終実行","結果","操作"])
        self.table.setColumnWidth(0,150); self.table.setColumnWidth(1,70)
        self.table.setColumnWidth(2,150); self.table.setColumnWidth(3,150)
        self.table.setColumnWidth(4,70);  self.table.setColumnWidth(5,260)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)
        self.load_tasks()

    @staticmethod
    def _fmt(raw):
        if not raw or raw.strip() == "": return "－"
        if raw.startswith("1999-11-30"): return "未実行"
        try:
            s = raw.split("+")[0].split(".")[0].strip()
            return datetime.fromisoformat(s).strftime("%Y-%m-%d %H:%M")
        except: return raw

    def load_tasks(self):
        self.table.setRowCount(0)
        for row, t in enumerate(get_simekiri_tasks()):
            name = t["name"]
            if name.startswith(TASK_BASE_NAME+"_"):
                cp = get_task_config_path(name[len(TASK_BASE_NAME)+1:])
                if os.path.exists(cp):
                    try:
                        with open(cp,"r",encoding="utf-8") as f:
                            name = json.load(f).get("title",name)
                    except: pass
            self.table.insertRow(row)
            status = "🔴 無効" if not t["enabled"] else ("🟢 有効" if t["state"]==3 else "⚪ 実行中")
            for col, val in enumerate([name, status,
                    self._fmt(t["next_run"]), self._fmt(t["last_run"]),
                    str(t["last_result"])]):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, col, item)

            bw = QWidget(); bl = QHBoxLayout(bw)
            bl.setContentsMargins(6,3,6,3); bl.setSpacing(6)
            eb = QPushButton("編集"); eb.setFixedWidth(60)
            db = QPushButton("削除"); db.setFixedWidth(60); db.setObjectName("btn_danger")
            rb = QPushButton("▶ 今すぐ実行")
            eb.clicked.connect(partial(self.edit_task,   t["name"]))
            db.clicked.connect(partial(self.delete_task, t["name"]))
            rb.clicked.connect(partial(self.run_task,    t["name"]))
            for b in (eb, db, rb): bl.addWidget(b)
            self.table.setCellWidget(row, 5, bw)
        self.table.resizeRowsToContents()

    def delete_task(self, tn):
        if QMessageBox.question(self,"削除確認",f"「{tn}」を削除しますか？",
                QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        did = tn.replace(TASK_BASE_NAME+"_",""); cp = get_task_config_path(did)
        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(
                None,"runas",sys.executable,
                f'"{os.path.abspath(__file__)}" --delete "{cp}"',None,1); return
        try:
            import win32com.client
            s = win32com.client.Dispatch("Schedule.Service"); s.Connect()
            s.GetFolder("\\").DeleteTask(tn,0)
            if os.path.exists(cp): os.remove(cp)
            QMessageBox.information(self,"削除完了","削除しました")
        except Exception as e: QMessageBox.warning(self,"削除失敗",str(e))
        self.load_tasks()

    def run_task(self, tn):
        cp = get_task_config_path(tn.replace(TASK_BASE_NAME+"_",""))
        if not os.path.exists(cp):
            QMessageBox.warning(self,"エラー","設定ファイルが見つかりません"); return
        import threading
        def _r():
            try:
                res = simekiri_notify.run_notify(cp, test_mode=False)
                msg = "通知を送信しました ✅" if res==0 else "エラーが発生しました。ログを確認してください。"
                QTimer.singleShot(0, lambda: (
                    QMessageBox.information(self,"完了",msg) if res==0
                    else QMessageBox.warning(self,"エラー",msg)
                ))
            except Exception as e:
                QTimer.singleShot(0, lambda: QMessageBox.warning(self,"エラー",str(e)))
        threading.Thread(target=_r, daemon=True).start()
        QMessageBox.information(self,"送信開始","送信中です…\n完了後に結果が表示されます。")

    def edit_task(self, tn):
        did = tn.replace(TASK_BASE_NAME+"_",""); cp = get_task_config_path(did)
        if not os.path.exists(cp):
            QMessageBox.warning(self,"エラー","設定ファイルが見つかりません"); return
        with open(cp,"r",encoding="utf-8") as f: cfg = json.load(f)
        dlg = TaskEditDialog(cfg, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            with open(cp,"r",encoding="utf-8") as f: upd = json.load(f)
            if self.main_app: self.main_app.update_task(upd)
            QMessageBox.information(self,"保存完了",f"「{upd.get('title','')}」を更新しました")
            self.load_tasks()


# ===================================================
# タスク編集ダイアログ
# ===================================================

class TaskEditDialog(GoogleAuthMixin, QDialog):
    def __init__(self, cfg, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.setWindowTitle(f"編集 – {cfg.get('title','')}")
        self.resize(560, 720)

        dark = self.palette().color(QPalette.ColorRole.Window).lightness() < 128
        self.setStyleSheet(make_stylesheet(dark))

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        inner  = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(20,16,20,24); layout.setSpacing(2)
        scroll.setWidget(inner)
        outer = QVBoxLayout(self); outer.setContentsMargins(0,0,0,0); outer.addWidget(scroll)

        # データソース
        layout.addWidget(section_header("データソース"))
        layout.addLayout(field_row("読み込み元", "datasource"))
        self.source_combo = QComboBox()
        self.source_combo.addItems(["Excel ファイル","Google スプレッドシート"])
        self.source_combo.setCurrentIndex(1 if cfg.get("data_source")=="sheets" else 0)
        self.source_combo.currentIndexChanged.connect(self._on_source_changed)
        layout.addWidget(self.source_combo)

        self.excel_group = QWidget()
        evl = QVBoxLayout(self.excel_group); evl.setContentsMargins(0,4,0,0); evl.setSpacing(4)
        evl.addLayout(field_row("Excel ファイルのパス"))
        ehl = QHBoxLayout(); ehl.setSpacing(6)
        self.excel_input = QLineEdit(cfg.get("excel_path",""))
        bb = QPushButton("参照"); bb.setFixedWidth(64); bb.clicked.connect(self.browse_excel)
        ehl.addWidget(self.excel_input); ehl.addWidget(bb); evl.addLayout(ehl)
        layout.addWidget(self.excel_group)

        self.sheets_group = QWidget()
        svl = QVBoxLayout(self.sheets_group); svl.setContentsMargins(0,4,0,0); svl.setSpacing(4)
        svl.addLayout(field_row("スプレッドシート URL","google_auth"))
        self.sheets_url_input = QLineEdit(cfg.get("sheets_url",""))
        svl.addWidget(self.sheets_url_input)
        ah = QHBoxLayout(); ah.setSpacing(10)
        self.auth_btn = QPushButton("🔐  Google で認可する")
        self.auth_btn.setObjectName("btn_primary")
        self.auth_btn.clicked.connect(self._authorize_google)
        self.auth_status_label = QLabel()
        ah.addWidget(self.auth_btn); ah.addWidget(self.auth_status_label); ah.addStretch()
        svl.addLayout(ah)
        layout.addWidget(self.sheets_group)
        self._on_source_changed(self.source_combo.currentIndex())
        self._refresh_auth_label()

        # 通知先
        layout.addWidget(section_header("通知先"))
        layout.addLayout(field_row("Webhook URL","webhook"))
        self.webhook_input = QLineEdit(cfg.get("webhook_url",""))
        layout.addWidget(self.webhook_input)

        layout.addLayout(field_row("締切何日前に通知","days_before"))
        self.days_spin = QSpinBox(); self.days_spin.setRange(0,60)
        self.days_spin.setValue(cfg.get("days_before_deadline",3))
        self.days_spin.setFixedWidth(100)
        layout.addWidget(self.days_spin)

        # メンション
        layout.addLayout(field_row("担当者メンション","mention"))
        self.mention_checkbox = QCheckBox("メンションを有効にする")
        self.mention_checkbox.setChecked(cfg.get("mention_enabled",False))
        layout.addWidget(self.mention_checkbox)
        self.mention_box = QWidget()
        self.mention_layout = QVBoxLayout(self.mention_box)
        self.mention_layout.setContentsMargins(20,0,0,0); self.mention_layout.setSpacing(2)
        for m in (cfg.get("mentions") or [{"name":"","id":""}]):
            r = RowInput("担当名","ユーザーID",self.mention_layout, deletable=bool(cfg.get("mentions")))
            r.short.setText(m.get("name","")); r.long.setText(m.get("id",""))
            self.mention_layout.addWidget(r)
        for i in range(self.mention_layout.count()):
            w = self.mention_layout.itemAt(i).widget()
            if isinstance(w,RowInput): w.update_delete_state()
        layout.addWidget(self.mention_box)

        # 確認待ち通知
        layout.addWidget(section_header("確認待ち通知"))
        layout.addLayout(field_row("レビュアーへの通知","reviewer"))
        self.reviewer_checkbox = QCheckBox("確認待ちタスクを別の担当者に通知する")
        self.reviewer_checkbox.setChecked(cfg.get("reviewer_enabled",False))
        self.reviewer_checkbox.stateChanged.connect(self._on_reviewer_toggled)
        layout.addWidget(self.reviewer_checkbox)

        self.reviewer_group = QWidget()
        rvl = QVBoxLayout(self.reviewer_group); rvl.setContentsMargins(20,4,0,0); rvl.setSpacing(4)
        rvl.addLayout(field_row("レビュアー Webhook URL（省略可）"))
        self.reviewer_webhook_input = QLineEdit(cfg.get("reviewer_webhook_url",""))
        rvl.addWidget(self.reviewer_webhook_input)
        rvl.addLayout(field_row("レビュアーのメンション"))
        self.reviewer_mention_layout = QVBoxLayout(); self.reviewer_mention_layout.setSpacing(2)
        for m in (cfg.get("reviewer_mentions") or [{"name":"","id":""}]):
            r = RowInput("担当名","レビュアーID",self.reviewer_mention_layout, deletable=bool(cfg.get("reviewer_mentions")))
            r.short.setText(m.get("name","")); r.long.setText(m.get("id",""))
            self.reviewer_mention_layout.addWidget(r)
        rvm = QWidget(); rvm.setLayout(self.reviewer_mention_layout)
        rvl.addWidget(rvm)
        layout.addWidget(self.reviewer_group)
        self.reviewer_group.setVisible(cfg.get("reviewer_enabled",False))

        # 自動連絡
        layout.addWidget(section_header("自動連絡"))
        layout.addLayout(field_row("タスクスケジューラ連携","auto_notify"))
        self.auto_checkbox = QCheckBox("自動送信を有効にする")
        self.auto_checkbox.setChecked(cfg.get("auto_notify",False))
        layout.addWidget(self.auto_checkbox)

        tl = QHBoxLayout(); tl.setSpacing(10)
        tl.addWidget(QLabel("送信時刻"))
        self.time_edit = QTimeEdit(QTime.fromString(cfg.get("notify_time","09:00"),"HH:mm"))
        self.time_edit.setFixedWidth(100)
        tl.addWidget(self.time_edit)
        tl.addSpacing(20); tl.addWidget(QLabel("頻度（日おき）"))
        self.interval_spin = QSpinBox(); self.interval_spin.setRange(1,30)
        self.interval_spin.setValue(cfg.get("notify_interval_days",1))
        self.interval_spin.setFixedWidth(80)
        tl.addWidget(self.interval_spin); tl.addStretch()
        layout.addLayout(tl)

        layout.addLayout(field_row("制作期間"))
        dl = QHBoxLayout(); dl.setSpacing(8)
        def pd(s, fb):
            d = QDate.fromString(s,"yyyy-MM-dd"); return d if d.isValid() else fb
        self.start_date_edit = QDateEdit(pd(cfg.get("start_date",""),QDate.currentDate()))
        self.start_date_edit.setCalendarPopup(True)
        self.end_date_edit   = QDateEdit(pd(cfg.get("end_date",""),QDate.currentDate().addYears(1)))
        self.end_date_edit.setCalendarPopup(True)
        dl.addWidget(QLabel("開始")); dl.addWidget(self.start_date_edit)
        dl.addSpacing(8); dl.addWidget(QLabel("終了")); dl.addWidget(self.end_date_edit); dl.addStretch()
        layout.addLayout(dl)

        layout.addSpacing(16)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine); layout.addWidget(sep)
        layout.addSpacing(10)
        sb = QPushButton("変更を保存する"); sb.setObjectName("btn_primary")
        sb.clicked.connect(self.save); layout.addWidget(sb)
        layout.addStretch()

    def _on_source_changed(self, i):
        self.excel_group.setVisible(i == 0); self.sheets_group.setVisible(i == 1)

    def _on_reviewer_toggled(self, state):
        self.reviewer_group.setVisible(state == Qt.CheckState.Checked.value)

    def browse_excel(self):
        p, _ = QFileDialog.getOpenFileName(self,"Excel を選択","","Excel (*.xlsx *.xls)")
        if p: self.excel_input.setText(p)

    def _collect_mentions(self, lyt):
        res = []
        for i in range(lyt.count()):
            w = lyt.itemAt(i).widget()
            if isinstance(w,RowInput):
                s,l = w.get()
                if s or l: res.append({"name":s,"id":l})
        return res

    def save(self):
        is_sheets = (self.source_combo.currentIndex() == 1)
        self.cfg.update({
            "data_source":          "sheets" if is_sheets else "excel",
            "excel_path":           self.excel_input.text() if not is_sheets else "",
            "sheets_url":           self.sheets_url_input.text() if is_sheets else "",
            "webhook_url":          self.webhook_input.text(),
            "days_before_deadline": self.days_spin.value(),
            "mention_enabled":      self.mention_checkbox.isChecked(),
            "mentions":             self._collect_mentions(self.mention_layout),
            "reviewer_enabled":     self.reviewer_checkbox.isChecked(),
            "reviewer_webhook_url": self.reviewer_webhook_input.text(),
            "reviewer_mentions":    self._collect_mentions(self.reviewer_mention_layout),
            "auto_notify":          self.auto_checkbox.isChecked(),
            "notify_time":          self.time_edit.time().toString("HH:mm"),
            "notify_interval_days": self.interval_spin.value(),
            "start_date":           self.start_date_edit.date().toString("yyyy-MM-dd"),
            "end_date":             self.end_date_edit.date().toString("yyyy-MM-dd"),
        })
        with open(get_config_path(self.cfg["deadline_id"]),"w",encoding="utf-8") as f:
            json.dump(self.cfg,f,ensure_ascii=False,indent=2)
        self.accept()


# ===================================================
# エントリポイント
# ===================================================

if __name__ == "__main__":
    try:
        if "--notify" in sys.argv:
            idx = sys.argv.index("--notify") + 1
            sys.exit(simekiri_notify.run_notify(sys.argv[idx] if len(sys.argv)>idx else None))

        if ADMIN_FLAG in sys.argv or "--delete" in sys.argv:
            fl  = "--delete" if "--delete" in sys.argv else ADMIN_FLAG
            idx = sys.argv.index(fl) + 1
            cp  = sys.argv[idx] if len(sys.argv)>idx else None
            if not cp or not os.path.exists(cp): print("Config missing"); sys.exit(1)
            with open(cp,"r",encoding="utf-8") as f: cfg = json.load(f)
            tn = get_task_name(cfg["deadline_id"])
            import win32com.client
            s = win32com.client.Dispatch("Schedule.Service"); s.Connect()
            root = s.GetFolder("\\")
            if fl == "--delete":
                try: root.DeleteTask(tn,0); print(f"削除: {tn}")
                except Exception as e: print("失敗:",e)
                if os.path.exists(cp): os.remove(cp)
            else:
                register_task_admin(cfg)
            sys.exit(0)

        app = QApplication(sys.argv)
        win = NotifierApp()
        win.show()
        sys.exit(app.exec())

    except Exception:
        with open(ERROR_LOG,"w",encoding="utf-8") as f:
            f.write(traceback.format_exc())
