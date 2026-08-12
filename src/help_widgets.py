#help_widgets.py
#
# フィールドごとのコンテキストヘルプ（「?」ボタン）とレイアウトヘルパー

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog, QToolButton, QLabel, QVBoxLayout, QHBoxLayout,
    QFrame, QTextBrowser, QPushButton,
)

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
        "4. credentials.json をダウンロードし、\n"
        "   「credentials.json を配置」ボタンで選択するか、\n"
        "   直接以下に配置してください：\n"
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
