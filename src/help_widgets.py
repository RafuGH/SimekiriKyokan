#help_widgets.py
#
# フィールドごとのコンテキストヘルプ（「?」ボタン）とレイアウトヘルパー
#
# 「?」を押すと、説明書（assets/manual/pageN.png）の該当ページを表示する。
# 画像が見つからない場合は、下の HELP_TEXTS の文章で説明する。

import os
import sys

from PyQt6.QtCore import Qt, QEvent, QTimer, QUrl, pyqtSignal
from PyQt6.QtGui import QDesktopServices, QPixmap
from PyQt6.QtWidgets import (
    QDialog, QToolButton, QLabel, QVBoxLayout, QHBoxLayout,
    QFrame, QTextBrowser, QPushButton, QScrollArea, QStackedWidget,
)

MANUAL_PAGE_COUNT = 8

# 説明書ページ画像の座標系（CSS px）。上端のタブ列の位置はクリック判定に使う。
_PAGE_W, _PAGE_H = 794, 1123
_TAB_LEFT, _TAB_BOTTOM = 262, 42

# 「?」の種類 → (説明書のページ, そのページ内で最初に見せたい位置のY座標)
HELP_PAGES = {
    "title":       (1, 55),
    "category":    (1, 185),
    "datasource":  (1, 460),
    "webhook":     (1, 665),
    "days_before": (1, 808),
    "mention":     (1, 912),
    "google_auth": (4, 0),
    "reviewer":    (2, 0),
    "submission":  (2, 530),
    "auto_notify": (3, 0),
}


def _resource_dirs():
    """インストール版は exe と同じフォルダ、開発時はリポジトリの assets を探す"""
    dirs = []
    if getattr(sys, "frozen", False):
        dirs.append(os.path.dirname(sys.executable))
    here = os.path.dirname(os.path.abspath(__file__))
    dirs += [here, os.path.normpath(os.path.join(here, "..", "assets"))]
    return dirs


def manual_page_path(page: int):
    for d in _resource_dirs():
        path = os.path.join(d, "manual", f"page{page}.png")
        if os.path.exists(path):
            return path
    return None


def manual_pdf_path():
    for d in _resource_dirs():
        path = os.path.join(d, "SimekiriKyokan_Manual.pdf")
        if os.path.exists(path):
            return path
    return None


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
        "「Google で認可する」ボタンで事前に認証が必要です。\n"
        "認証後は「📂 Drive から選択」ボタンで、Drive 上の\n"
        "スプレッドシートを一覧から選べます（URL の貼り付けは不要）。"
    ),
    "google_auth": (
        "Google 認証について",
        "Google Sheets を使うには一度だけ認証が必要です。\n\n"
        "【手順】\n"
        "1. Google Cloud Console でプロジェクトを作成\n"
        "2. Google Sheets API と Google Drive API を有効化\n"
        "3. OAuth 2.0 クライアントIDを作成\n"
        "   （種類: ウェブアプリケーション、\n"
        "     リダイレクト URI: http://localhost:8080/callback）\n"
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
    "submission": (
        "提出フォルダ監視",
        "指定したフォルダに新しいファイルが置かれたとき、または\n"
        "既存ファイルが更新されたときに Webhook で通知します。\n\n"
        "【監視する場所】\n"
        "・ローカル / OneDrive の同期フォルダ\n"
        "・Google Drive のフォルダ（「Drive から選択」で指定）\n"
        "のどちらかを選べます。\n\n"
        "【OneDrive / Teams で使う場合】\n"
        "OneDrive や Teams の「ファイル」タブの中身は、PC上では\n"
        "同期フォルダとして見えています。そのフォルダを指定してください。\n"
        "例：C:\\Users\\名前\\OneDrive - 会社名\\チーム\\提出\n\n"
        "メンバーが Excel Online 等でファイルを提出・更新すると、\n"
        "同期後のチェックタイミングで検知されます。\n\n"
        "【チェックのタイミング】\n"
        "「自動連絡」で登録したスケジュール実行時と、\n"
        "管理画面の「今すぐ実行」時に確認されます。\n\n"
        "【初回について】\n"
        "有効化直後の初回スキャンでは通知しません。\n"
        "（既存ファイルが全件通知されてしまうのを防ぐため）\n"
        "次回以降、変化があった分だけ通知します。\n\n"
        "【PCを起動していなくても即時通知したい場合】\n"
        "Microsoft の Power Automate で\n"
        "「ファイルが作成されたとき（OneDrive/SharePoint）」→\n"
        "「HTTP」アクションで Webhook URL に POST するフローを作ると、\n"
        "クラウド側で即座に通知できます。併用も可能です。"
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
        page, anchor = HELP_PAGES.get(self.help_key, (1, 0))
        dlg = HelpDialog(title, body, self.window(), page=page, anchor=anchor)
        dlg.exec()


class _PageImage(QLabel):
    """説明書ページの画像。上端のタブ列をクリックするとそのページへ移る。"""
    tab_clicked = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.setMouseTracking(True)
        self.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)

    def _tab_at(self, pos):
        pm = self.pixmap()
        if pm is None or pm.isNull():
            return None
        shown_w = pm.width() / pm.devicePixelRatio()
        left = (self.width() - shown_w) / 2
        x = (pos.x() - left) * _PAGE_W / shown_w
        y = pos.y() * _PAGE_W / shown_w
        if y > _TAB_BOTTOM or x < _TAB_LEFT or x > _PAGE_W:
            return None
        tab_w = (_PAGE_W - _TAB_LEFT) / MANUAL_PAGE_COUNT
        return min(MANUAL_PAGE_COUNT, int((x - _TAB_LEFT) // tab_w) + 1)

    def mouseMoveEvent(self, e):
        on_tab = self._tab_at(e.position()) is not None
        self.setCursor(Qt.CursorShape.PointingHandCursor if on_tab else Qt.CursorShape.ArrowCursor)
        super().mouseMoveEvent(e)

    def mousePressEvent(self, e):
        page = self._tab_at(e.position())
        if page is not None:
            self.tab_clicked.emit(page)
            return
        super().mousePressEvent(e)


class HelpDialog(QDialog):
    """
    ヘルプ表示ダイアログ。
    説明書のページ画像があれば、該当ページの該当箇所を表示する（ページ移動も可）。
    無ければ従来どおり文章で説明する。
    """
    def __init__(self, title: str = "説明書", body: str = "", parent=None, page: int = 1, anchor: int = 0):
        super().__init__(parent)
        self.setWindowTitle(f"ヘルプ – {title}")
        self.setWindowFlags(
            Qt.WindowType.Dialog |
            Qt.WindowType.WindowCloseButtonHint |
            Qt.WindowType.WindowMaximizeButtonHint
        )
        self._pixmaps = {}
        self._page = page
        self._has_manual = manual_page_path(page) is not None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(10)

        hdr = QLabel(title)
        hdr.setStyleSheet("font-size:15px; font-weight:600;")
        layout.addWidget(hdr)

        self.stack = QStackedWidget()
        layout.addWidget(self.stack, 1)

        # 説明書（画像）
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        # 表示幅が変わるたびに（初回表示・リサイズ時）画像を幅に合わせ直す
        self.scroll.viewport().installEventFilter(self)
        self.image = _PageImage()
        self.image.tab_clicked.connect(self.show_page)
        self.scroll.setWidget(self.image)
        self.stack.addWidget(self.scroll)

        # 文章
        browser = QTextBrowser()
        browser.setPlainText(body or "説明がありません。")
        self.stack.addWidget(browser)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)

        row = QHBoxLayout(); row.setSpacing(6)
        self.prev_btn = QPushButton("◀")
        self.prev_btn.setToolTip("前のページ")
        self.prev_btn.clicked.connect(lambda: self.show_page(self._page - 1))
        self.page_lbl = QLabel()
        self.page_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_lbl.setMinimumWidth(56)
        self.next_btn = QPushButton("▶")
        self.next_btn.setToolTip("次のページ")
        self.next_btn.clicked.connect(lambda: self.show_page(self._page + 1))
        self.mode_btn = QPushButton("文章で見る")
        self.mode_btn.clicked.connect(self._toggle_mode)
        self.pdf_btn = QPushButton("PDFで開く")
        self.pdf_btn.clicked.connect(self._open_pdf)
        close_btn = QPushButton("閉じる")
        close_btn.setObjectName("btn_primary")
        close_btn.clicked.connect(self.accept)
        for w in (self.prev_btn, self.page_lbl, self.next_btn):
            row.addWidget(w)
        row.addStretch()
        for w in (self.mode_btn, self.pdf_btn, close_btn):
            row.addWidget(w)
        layout.addLayout(row)

        if self._has_manual:
            screen = self.screen().availableGeometry() if self.screen() else None
            height = min(900, int(screen.height() * 0.9)) if screen else 860
            self.resize(740, height)
            self.show_page(page)
            # レイアウト確定後に該当箇所までスクロールする
            QTimer.singleShot(0, lambda: self._scroll_to(anchor))
        else:
            # 説明書の画像が無い環境では従来どおり文章のみ
            self.setFixedWidth(420)
            self.stack.setCurrentIndex(1)
            browser.setMinimumHeight(120)
            browser.setMaximumHeight(320)
            for w in (self.prev_btn, self.page_lbl, self.next_btn, self.mode_btn):
                w.hide()
        self.pdf_btn.setVisible(manual_pdf_path() is not None)

    # ---- ページ表示 ----
    def show_page(self, page: int):
        page = max(1, min(MANUAL_PAGE_COUNT, page))
        path = manual_page_path(page)
        if path is None:
            return
        if page not in self._pixmaps:
            self._pixmaps[page] = QPixmap(path)
        self._page = page
        self._render()
        self.scroll.verticalScrollBar().setValue(0)
        self.page_lbl.setText(f"{page} / {MANUAL_PAGE_COUNT}")
        self.prev_btn.setEnabled(page > 1)
        self.next_btn.setEnabled(page < MANUAL_PAGE_COUNT)

    def _render(self):
        src = self._pixmaps.get(self._page)
        if src is None or src.isNull():
            return
        width = max(200, self.scroll.viewport().width())
        dpr = self.devicePixelRatioF()
        pm = src.scaledToWidth(int(width * dpr), Qt.TransformationMode.SmoothTransformation)
        pm.setDevicePixelRatio(dpr)
        self.image.setPixmap(pm)

    def _scroll_to(self, anchor: int):
        if anchor <= 0:
            return
        pm = self.image.pixmap()
        if pm is None or pm.isNull():
            return
        shown_h = pm.height() / pm.devicePixelRatio()
        self.scroll.verticalScrollBar().setValue(int(shown_h * anchor / _PAGE_H))

    def eventFilter(self, obj, e):
        if obj is self.scroll.viewport() and e.type() == QEvent.Type.Resize and self._has_manual:
            self._render()
        return super().eventFilter(obj, e)

    def _toggle_mode(self):
        to_text = self.stack.currentIndex() == 0
        self.stack.setCurrentIndex(1 if to_text else 0)
        self.mode_btn.setText("説明書で見る" if to_text else "文章で見る")
        for w in (self.prev_btn, self.page_lbl, self.next_btn):
            w.setVisible(not to_text)

    def _open_pdf(self):
        path = manual_pdf_path()
        if path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))


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
