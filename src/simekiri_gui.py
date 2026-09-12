#simekiri_gui.py

import sys, os, json, shutil, traceback
import threading
import webbrowser

from PyQt6.QtWidgets import (
    QApplication, QWidget, QScrollArea,
    QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QComboBox,
    QSpinBox, QCheckBox, QDateEdit, QTimeEdit, QFileDialog, QMessageBox,
    QSizePolicy, QFrame, QProgressDialog,
)
from PyQt6.QtCore import QTime, QDate, Qt, pyqtSignal

import simekiri_notify
import google_auth_helper
import app_settings
import update_checker
import drive_picker
import google_drive
from theme import apply_theme
from help_widgets import field_row, section_header
from account_badge import AccountBadge
from google_auth_mixin import GoogleAuthMixin
from task_scheduler import (
    APP_DIR, is_admin, get_config_path, get_task_name, generate_deadline_id,
    register_task_admin, relaunch_as_admin, ADMIN_FLAG,
)
from widgets import RowInput
from task_manager_window import TaskManagerWindow


# ===================================================
# メインGUI
# ===================================================

class NotifierApp(GoogleAuthMixin, QWidget):
    # アップデート確認はバックグラウンドスレッドで行うため、
    # 結果はシグナル経由でGUIスレッドに渡す
    update_found = pyqtSignal(dict)
    download_progress = pyqtSignal(int, int)
    download_finished = pyqtSignal(str)
    download_failed = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("締切教官")
        self.setMinimumWidth(560)
        self.resize(580, 880)

        apply_theme(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---- タイトルバー ----
        titlebar = QWidget(); titlebar.setObjectName("titlebar")
        tb_layout = QHBoxLayout(titlebar)
        tb_layout.setContentsMargins(16, 10, 12, 10)
        app_lbl = QLabel("締切教官")
        app_lbl.setStyleSheet("font-size:15px; font-weight:700; letter-spacing:0.5px;")
        version_lbl = QLabel(f"v{update_checker.APP_VERSION}")
        version_lbl.setObjectName("desc_lbl")
        self.manual_btn = QPushButton("📖 マニュアル")
        self.manual_btn.clicked.connect(self.open_manual)
        self.theme_btn = QPushButton()
        self.theme_btn.setFixedWidth(46)
        self.theme_btn.setToolTip("ライト / ダークテーマを切り替え")
        self.theme_btn.clicked.connect(self._toggle_theme)
        self._update_theme_btn()
        self.account_badge = AccountBadge()
        tb_layout.addWidget(app_lbl)
        tb_layout.addWidget(version_lbl)
        tb_layout.addStretch()
        tb_layout.addWidget(self.manual_btn)
        tb_layout.addWidget(self.theme_btn)
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
        sheet_hl = QHBoxLayout(); sheet_hl.setSpacing(6)
        self.sheets_url_input = QLineEdit()
        self.sheets_url_input.setPlaceholderText("https://docs.google.com/spreadsheets/d/...")
        self.pick_sheet_btn = QPushButton("📂 Drive から選択")
        self.pick_sheet_btn.setFixedWidth(150)
        self.pick_sheet_btn.clicked.connect(self.pick_spreadsheet_from_drive)
        sheet_hl.addWidget(self.sheets_url_input)
        sheet_hl.addWidget(self.pick_sheet_btn)
        svl.addLayout(sheet_hl)

        # credentials.json 配置ボタン＋状態表示
        cred_hl = QHBoxLayout(); cred_hl.setSpacing(10)
        self.cred_btn = QPushButton("📄 credentials.json を配置")
        self.cred_btn.clicked.connect(self._choose_credentials_file)
        self.cred_status_label = QLabel()
        cred_hl.addWidget(self.cred_btn)
        cred_hl.addWidget(self.cred_status_label)
        cred_hl.addStretch()
        svl.addLayout(cred_hl)

        # Google 認可ボタン＋状態表示
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
        self._init_google_auth()
        self._update_credentials_status()
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
        ml = QVBoxLayout(mb); ml.setContentsMargins(20, 0, 0, 0); ml.setSpacing(2)
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

        # ━━ 提出フォルダ監視 ━━
        layout.addWidget(section_header("提出フォルダ監視"))
        layout.addLayout(field_row("フォルダへの提出を通知（任意）", "submission"))
        self.submission_checkbox = QCheckBox("提出フォルダを監視して通知する")
        layout.addWidget(self.submission_checkbox)
        self.submission_checkbox.stateChanged.connect(self._on_submission_toggled)

        self.submission_group = QWidget()
        subvl = QVBoxLayout(self.submission_group)
        subvl.setContentsMargins(20, 4, 0, 0); subvl.setSpacing(4)

        subvl.addLayout(field_row("監視する場所"))
        self.submission_source_combo = QComboBox()
        self.submission_source_combo.addItems(["ローカル / OneDrive の同期フォルダ", "Google Drive のフォルダ"])
        self.submission_source_combo.currentIndexChanged.connect(self._on_submission_source_changed)
        subvl.addWidget(self.submission_source_combo)

        # ローカルフォルダ
        self.submission_local_group = QWidget()
        local_vl = QVBoxLayout(self.submission_local_group)
        local_vl.setContentsMargins(0, 0, 0, 0); local_vl.setSpacing(4)
        local_vl.addLayout(field_row("監視するフォルダ"))
        sub_hl = QHBoxLayout(); sub_hl.setSpacing(6)
        self.submission_folder_input = QLineEdit()
        self.submission_folder_input.setPlaceholderText("例：C:\\Users\\名前\\OneDrive - 会社名\\チーム\\提出")
        btn_sub = QPushButton("参照"); btn_sub.setFixedWidth(64)
        btn_sub.clicked.connect(self.browse_submission_folder)
        sub_hl.addWidget(self.submission_folder_input); sub_hl.addWidget(btn_sub)
        local_vl.addLayout(sub_hl)
        self.submission_recursive_checkbox = QCheckBox("サブフォルダも対象にする")
        self.submission_recursive_checkbox.setChecked(True)
        local_vl.addWidget(self.submission_recursive_checkbox)
        subvl.addWidget(self.submission_local_group)

        # Google Drive フォルダ
        self.submission_drive_group = QWidget()
        drive_vl = QVBoxLayout(self.submission_drive_group)
        drive_vl.setContentsMargins(0, 0, 0, 0); drive_vl.setSpacing(4)
        drive_vl.addLayout(field_row("Google Drive のフォルダ"))
        drive_hl = QHBoxLayout(); drive_hl.setSpacing(6)
        self.submission_drive_input = QLineEdit()
        self.submission_drive_input.setPlaceholderText("フォルダURL、または「Drive から選択」")
        self.pick_drive_folder_btn = QPushButton("📂 Drive から選択")
        self.pick_drive_folder_btn.setFixedWidth(150)
        self.pick_drive_folder_btn.clicked.connect(self.pick_submission_folder_from_drive)
        drive_hl.addWidget(self.submission_drive_input)
        drive_hl.addWidget(self.pick_drive_folder_btn)
        drive_vl.addLayout(drive_hl)
        subvl.addWidget(self.submission_drive_group)
        self.submission_drive_group.setVisible(False)
        self._submission_drive_folder_name = ""
        subvl.addLayout(field_row("対象の拡張子（省略可・カンマ区切り）"))
        self.submission_ext_input = QLineEdit()
        self.submission_ext_input.setPlaceholderText("例：.xlsx,.docx,.pdf（空欄ならすべて）")
        subvl.addWidget(self.submission_ext_input)
        subvl.addLayout(field_row("提出通知先 Webhook URL（省略可）"))
        self.submission_webhook_input = QLineEdit()
        self.submission_webhook_input.setPlaceholderText("省略すると上の URL を使用")
        subvl.addWidget(self.submission_webhook_input)
        layout.addWidget(self.submission_group)
        self.submission_group.setVisible(False)

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

        self.update_found.connect(self._on_update_found)
        self.download_progress.connect(self._on_download_progress)
        self.download_finished.connect(self._on_download_finished)
        self.download_failed.connect(self._on_download_failed)
        self._update_dialog = None
        self._start_update_check()

    # ---- アップデート確認 ----
    def _start_update_check(self):
        """起動時に新しいリリースが無いかバックグラウンドで確認する。"""
        if not app_settings.load_settings().get("check_update_on_launch", True):
            return

        def _check():
            try:
                release = update_checker.check_for_update()
            except Exception:
                return   # 確認できなくてもアプリの動作には影響させない
            if not release:
                return
            # 同じバージョンを毎回告知しない
            if release["version"] == app_settings.load_settings().get("last_notified_version"):
                return
            self.update_found.emit(release)

        threading.Thread(target=_check, daemon=True).start()

    def _on_update_found(self, release: dict):
        app_settings.save_settings({"last_notified_version": release.get("version", "")})
        reply = QMessageBox.question(
            self, "アップデートのお知らせ",
            update_checker.format_message(release),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        asset = update_checker.find_installer_asset(release)
        if not asset:
            # インストーラーが添付されていないリリースはページを開くだけにする
            webbrowser.open(release.get("url", update_checker.RELEASES_PAGE))
            return

        self._start_update_download(release, asset)

    def _start_update_download(self, release: dict, asset: dict):
        """インストーラーをダウンロードし、完了したら起動する。"""
        self._update_dialog = QProgressDialog(
            f"{release.get('version', '')} をダウンロードしています…", "キャンセル", 0, 100, self
        )
        self._update_dialog.setWindowTitle("アップデート")
        self._update_dialog.setAutoClose(False)
        self._update_dialog.setAutoReset(False)
        self._update_dialog.setMinimumDuration(0)
        self._update_dialog.setValue(0)

        def _download():
            try:
                path = update_checker.download_installer(
                    asset,
                    progress_cb=lambda received, total: self.download_progress.emit(received, total),
                )
            except Exception as e:
                self.download_failed.emit(str(e))
                return
            self.download_finished.emit(path)

        threading.Thread(target=_download, daemon=True).start()

    def _on_download_progress(self, received: int, total: int):
        if not self._update_dialog:
            return
        if self._update_dialog.wasCanceled():
            return
        if total > 0:
            self._update_dialog.setValue(int(received * 100 / total))
            self._update_dialog.setLabelText(
                f"ダウンロード中… {received // (1024 * 1024)} MB / {total // (1024 * 1024)} MB"
            )
        else:
            self._update_dialog.setLabelText(f"ダウンロード中… {received // (1024 * 1024)} MB")

    def _on_download_finished(self, installer_path: str):
        canceled = bool(self._update_dialog and self._update_dialog.wasCanceled())
        if self._update_dialog:
            self._update_dialog.close()
            self._update_dialog = None
        if canceled:
            return

        QMessageBox.information(
            self, "アップデート",
            "ダウンロードが完了しました。\n"
            "インストーラーを起動して更新します。\n"
            "（更新のため、締切教官はいったん終了します）"
        )
        try:
            started = update_checker.launch_installer(installer_path)
        except Exception as e:
            QMessageBox.warning(self, "アップデート", f"インストーラーを起動できませんでした:\n{e}")
            return
        if not started:
            QMessageBox.warning(
                self, "アップデート",
                "インストーラーの起動がキャンセルされました。\n"
                f"手動で実行してください:\n{installer_path}"
            )
            return
        QApplication.quit()

    def _on_download_failed(self, message: str):
        if self._update_dialog:
            self._update_dialog.close()
            self._update_dialog = None
        QMessageBox.warning(
            self, "アップデート",
            f"ダウンロードに失敗しました:\n{message}\n\nダウンロードページを開きます。"
        )
        webbrowser.open(update_checker.RELEASES_PAGE)

    # ---- テーマ切り替え ----
    def _update_theme_btn(self):
        # 次に切り替わる先のアイコンを表示する
        self.theme_btn.setText("☀" if app_settings.is_dark() else "🌙")

    def _toggle_theme(self):
        app_settings.set_theme("light" if app_settings.is_dark() else "dark")
        apply_theme(self)
        self._update_theme_btn()
        # 開いている管理ウィンドウにも即座に反映する
        win = getattr(self, "task_list_window", None)
        if win is not None:
            try:
                apply_theme(win)
            except RuntimeError:
                pass   # 既に閉じられている場合

    # ---- Google 連携 ----
    def _choose_credentials_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "credentials.json を選択", "", "JSON (*.json)")
        if not path:
            return
        try:
            google_auth_helper.set_credentials_file(path)
            QMessageBox.information(self, "配置完了", "credentials.json を配置しました")
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"credentials.json の配置に失敗しました:\n{e}")
        self._update_credentials_status()

    def _update_credentials_status(self):
        if google_auth_helper.has_credentials():
            self.cred_status_label.setText("✅ 配置済み")
        else:
            self.cred_status_label.setText("❌ 未配置")

    # ---- ソース切り替え ----
    def _on_source_changed(self, index):
        is_sheets = (index == 1)
        self.excel_group.setVisible(not is_sheets)
        self.sheets_group.setVisible(is_sheets)

    def pick_spreadsheet_from_drive(self):
        picked = drive_picker.pick_spreadsheet(self)
        if picked:
            self.sheets_url_input.setText(picked["url"])

    def pick_submission_folder_from_drive(self):
        picked = drive_picker.pick_folder(self)
        if picked:
            self.submission_drive_input.setText(picked["url"])
            self._submission_drive_folder_name = picked.get("name", "")

    def _on_submission_source_changed(self, index):
        is_drive = (index == 1)
        self.submission_local_group.setVisible(not is_drive)
        self.submission_drive_group.setVisible(is_drive)

    def _on_submission_toggled(self, state):
        self.submission_group.setVisible(state == Qt.CheckState.Checked.value)

    def browse_submission_folder(self):
        path = QFileDialog.getExistingDirectory(self, "監視するフォルダを選択")
        if path:
            self.submission_folder_input.setText(path)

    def _on_reviewer_toggled(self, state):
        self.reviewer_group.setVisible(state == Qt.CheckState.Checked.value)

    def browse_excel(self):
        p, _ = QFileDialog.getOpenFileName(self, "Excelを選択", "", "Excel (*.xlsx *.xls)")
        if p:
            self.excel_input.setText(p)

    def open_manual(self):
        try:
            base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
            pdf_path = os.path.join(base_dir, "SimekiriKyokan_Manual.pdf")
            if os.path.exists(pdf_path):
                os.startfile(pdf_path)
            else:
                QMessageBox.warning(self, "エラー", "マニュアルPDFが見つかりません")
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"マニュアルを開けませんでした:\n{e}")

    def generate_excel(self):
        base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "Tasks.xlsx")
        if not os.path.exists(template_path):
            QMessageBox.critical(self, "エラー", f"テンプレート Excel が見つかりません:\n{template_path}")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "保存先を選択", "Tasks.xlsx", "Excel Files (*.xlsx)")
        if not save_path:
            return
        if os.path.exists(save_path):
            reply = QMessageBox.question(
                self, "上書き確認", f"既存のファイルが存在します。\n上書きしますか？\n\n{save_path}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        try:
            shutil.copyfile(template_path, save_path)
            QMessageBox.information(self, "完了", f"生成しました:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"生成に失敗しました:\n{str(e)}")

    def _collect_mentions(self, layout_widget) -> list:
        mentions = []
        for i in range(layout_widget.count()):
            w = layout_widget.itemAt(i).widget()
            if isinstance(w, RowInput):
                short, long_ = w.get()
                if short or long_:
                    mentions.append({"name": short, "id": long_})
        return mentions

    def save_config(self):
        title = self.title_input.text().strip()
        if not title:
            QMessageBox.warning(self, "入力エラー", "締切名を入力してください")
            return

        is_sheets = (self.source_combo.currentIndex() == 1)

        if is_sheets:
            if not self.sheets_url_input.text().strip():
                QMessageBox.warning(self, "入力エラー", "スプレッドシート URL を入力してください")
                return
            if not google_auth_helper.has_token():
                QMessageBox.warning(self, "入力エラー", "Google 認証を先に完了してください")
                return
        else:
            if not self.excel_input.text().strip():
                QMessageBox.warning(self, "入力エラー", "Excelファイルを指定してください")
                return

        if not self.webhook_input.text().strip():
            QMessageBox.warning(self, "入力エラー", "Webhook URLを入力してください")
            return

        if self.submission_checkbox.isChecked():
            if self.submission_source_combo.currentIndex() == 1:
                if not google_drive.extract_folder_id(self.submission_drive_input.text()):
                    QMessageBox.warning(self, "入力エラー", "監視する Google Drive のフォルダを指定してください")
                    return
                if not google_auth_helper.has_token():
                    QMessageBox.warning(self, "入力エラー", "Google Drive を使うには先に Google 認証を完了してください")
                    return
            else:
                folder = self.submission_folder_input.text().strip()
                if not folder:
                    QMessageBox.warning(self, "入力エラー", "監視する提出フォルダを指定してください")
                    return
                if not os.path.isdir(folder):
                    QMessageBox.warning(self, "入力エラー", f"提出フォルダが見つかりません:\n{folder}")
                    return

        category  = self.category_combo.currentText()
        end_date  = self.end_date.date().toString("yyyy-MM-dd")
        deadline_id = generate_deadline_id(category, end_date, title)

        cfg = {
            "deadline_id":          deadline_id,
            "title":                title,
            "category":             category,
            # データソース
            "data_source":          "sheets" if is_sheets else "excel",
            "excel_path":           self.excel_input.text() if not is_sheets else "",
            "sheets_url":           self.sheets_url_input.text() if is_sheets else "",
            # 通知設定
            "webhook_url":          self.webhook_input.text(),
            "days_before_deadline": self.days_spin.value(),
            "mention_enabled":      self.mention_checkbox.isChecked(),
            "mentions":             self._collect_mentions(self.mention_layout),
            # 確認待ち通知
            "reviewer_enabled":     self.reviewer_checkbox.isChecked(),
            "reviewer_webhook_url": self.reviewer_webhook_input.text(),
            "reviewer_mentions":    self._collect_mentions(self.reviewer_mention_layout),
            # 提出フォルダ監視
            "submission_watch_enabled": self.submission_checkbox.isChecked(),
            "submission_source":        "drive" if self.submission_source_combo.currentIndex() == 1 else "local",
            "submission_folder":        self.submission_folder_input.text().strip(),
            "submission_drive_folder_id":   google_drive.extract_folder_id(self.submission_drive_input.text()),
            "submission_drive_folder_name": self._submission_drive_folder_name,
            "submission_extensions":    self.submission_ext_input.text().strip(),
            "submission_webhook_url":   self.submission_webhook_input.text().strip(),
            "submission_recursive":     self.submission_recursive_checkbox.isChecked(),
            # 自動通知
            "auto_notify":          self.auto_checkbox.isChecked(),
            "notify_time":          self.time_edit.time().toString("HH:mm"),
            "notify_interval_days": self.interval_spin.value(),
            "start_date":           self.start_date.date().toString("yyyy-MM-dd"),
            "end_date":             end_date,
        }

        config_path = get_config_path(deadline_id)
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

        self.config = cfg

        if cfg["auto_notify"]:
            try:
                if is_admin():
                    register_task_admin(cfg)
                    QMessageBox.information(self, "保存完了", "設定を保存してタスクを登録しました")
                else:
                    reply = QMessageBox.question(
                        self, "管理者権限確認",
                        "タスク登録には管理者権限が必要です。昇格しますか？",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.Yes:
                        relaunch_as_admin(config_path, ADMIN_FLAG)
                        QMessageBox.information(self, "保存完了", "設定を保存しました（管理者認証待ち）")
                    else:
                        QMessageBox.information(self, "保存完了", "設定を保存しました（タスク未登録）")
            except Exception as e:
                QMessageBox.warning(self, "エラー", f"タスク登録に失敗しました: {e}")
        else:
            QMessageBox.information(self, "保存完了", "設定を保存しました")

        self._reset_form()

    def _reset_form(self):
        self.config = {}
        for w in (self.title_input, self.excel_input, self.sheets_url_input,
                  self.webhook_input, self.reviewer_webhook_input,
                  self.submission_folder_input, self.submission_ext_input,
                  self.submission_webhook_input, self.submission_drive_input):
            w.clear()
        self.days_spin.setValue(3)
        self.mention_checkbox.setChecked(False)
        self.reviewer_checkbox.setChecked(False)
        self.submission_checkbox.setChecked(False)
        self.submission_recursive_checkbox.setChecked(True)
        self.submission_source_combo.setCurrentIndex(0)
        self._submission_drive_folder_name = ""
        self.auto_checkbox.setChecked(False)
        self.time_edit.setTime(QTime(9, 0))
        self.interval_spin.setValue(1)
        self.source_combo.setCurrentIndex(0)
        for lyt in (self.mention_layout, self.reviewer_mention_layout):
            while lyt.count():
                w = lyt.takeAt(0).widget()
                if w:
                    w.setParent(None)
        self.mention_layout.addWidget(RowInput("担当名", "ユーザーID", self.mention_layout, False))
        self.reviewer_mention_layout.addWidget(RowInput("担当名", "レビュアーID", self.reviewer_mention_layout, False))

    def run_notify(self):
        files = [f for f in os.listdir(APP_DIR) if f.endswith(".json")]
        if not files:
            QMessageBox.warning(self, "エラー", "保存された設定がありません")
            return
        files.sort(key=lambda x: os.path.getmtime(os.path.join(APP_DIR, x)), reverse=True)
        config_path = os.path.join(APP_DIR, files[0])
        simekiri_notify.run_notify(config_path, test_mode=True)

    def open_task_list(self):
        self.task_list_window = TaskManagerWindow(self)
        self.task_list_window.show()

    def update_task(self, cfg):
        """
        タスク編集ダイアログでの保存後にタスク定義を作り直す。
        register_task_admin() は登録時に同名タスクを削除してから
        再作成するため、事前の削除処理は不要。
        """
        config_path = get_config_path(cfg["deadline_id"])
        if is_admin():
            try:
                register_task_admin(cfg)
            except Exception as e:
                QMessageBox.warning(self, "タスク更新失敗", str(e))
            return
        relaunch_as_admin(config_path, ADMIN_FLAG)


# ===================================================
# エントリポイント
# ===================================================

def _run_notify_cli(config_path):
    sys.exit(simekiri_notify.run_notify(config_path))


def _run_admin_task_cli(config_path, delete=False):
    if not config_path or not os.path.exists(config_path):
        print("Config path missing")
        sys.exit(1)

    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    task_name = get_task_name(cfg["deadline_id"])

    if delete:
        import win32com.client
        service = win32com.client.Dispatch("Schedule.Service")
        service.Connect()
        root = service.GetFolder("\\")
        try:
            root.DeleteTask(task_name, 0)
            print(f"削除成功: {task_name}")
        except Exception as e:
            print("削除失敗:", e)
        if os.path.exists(config_path):
            os.remove(config_path)
        sys.exit(0)
    else:
        register_task_admin(cfg)
        sys.exit(0)


if __name__ == "__main__":
    ERROR_LOG = os.path.join(APP_DIR, "gui_error_log.txt")
    try:
        if "--notify" in sys.argv:
            idx = sys.argv.index("--notify") + 1
            config_path = sys.argv[idx] if len(sys.argv) > idx else None
            _run_notify_cli(config_path)

        elif "--delete" in sys.argv:
            idx = sys.argv.index("--delete") + 1
            config_path = sys.argv[idx] if len(sys.argv) > idx else None
            _run_admin_task_cli(config_path, delete=True)

        elif ADMIN_FLAG in sys.argv:
            idx = sys.argv.index(ADMIN_FLAG) + 1
            config_path = sys.argv[idx] if len(sys.argv) > idx else None
            _run_admin_task_cli(config_path, delete=False)

        else:
            app = QApplication(sys.argv)
            win = NotifierApp()
            win.show()
            sys.exit(app.exec())

    except Exception:
        with open(ERROR_LOG, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
