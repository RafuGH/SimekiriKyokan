#simekiri_gui.py

import sys, os, json, shutil, traceback
import threading

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QCheckBox, QSpinBox, QTimeEdit, QDateEdit,
    QScrollArea, QFileDialog, QMessageBox,
)
from PyQt6.QtCore import QTime, QDate, Qt

import simekiri_notify
import google_auth_helper
from task_scheduler import (
    APP_DIR, is_admin, get_config_path, get_task_name, generate_deadline_id,
    task_exists, register_task_admin, relaunch_as_admin, ADMIN_FLAG,
)
from widgets import RowInput
from task_manager_window import TaskManagerWindow

APP_VERSION = "v2.1"


# ===================================================
# メインGUI
# ===================================================

class NotifierApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"締切教官 {APP_VERSION}")
        self.resize(560, 900)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        inner = QWidget()
        layout = QVBoxLayout(inner)
        scroll_area.setWidget(inner)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll_area)

        palette = self.palette()
        base_color = palette.color(palette.ColorRole.Window)
        is_dark = base_color.lightness() < 128

        # ---- タイトル行 ----
        title_row = QHBoxLayout()
        title_label = QLabel("締切名")
        self.help_btn = QPushButton("？")
        self.help_btn.setFixedSize(24, 24)
        if is_dark:
            self.help_btn.setStyleSheet("""
                QPushButton { font-size:14px; font-weight:bold;
                    border:1px solid palette(mid); border-radius:4px;
                    padding:0px; background-color: palette(button); color: palette(button-text); }
                QPushButton:hover { background-color: palette(light); }
            """)
        else:
            self.help_btn.setStyleSheet("""
                QPushButton { font-size:14px; font-weight:bold;
                    border:1px solid #cccccc; border-radius:4px;
                    padding:0px; background-color: #f5f5f5; color: black; }
                QPushButton:hover { background-color: #e0e0e0; }
            """)
        title_label.setContentsMargins(0, 20, 0, 0)
        self.help_btn.clicked.connect(self.open_manual)
        title_row.addWidget(title_label)
        title_row.addStretch()
        title_row.addWidget(self.help_btn)
        layout.addLayout(title_row)

        self.title_input = QLineEdit()
        layout.addWidget(self.title_input)

        # ---- カテゴリ ----
        layout.addWidget(QLabel("カテゴリ"))
        self.category_combo = QComboBox()
        self.category_combo.addItems(["report", "game", "school", "work", "personal"])
        layout.addWidget(self.category_combo)

        # ---- データソース切り替え ----
        layout.addWidget(QLabel("──────── データソース ────────"))
        self.source_combo = QComboBox()
        self.source_combo.addItems(["Excel ファイル", "Google スプレッドシート"])
        self.source_combo.currentIndexChanged.connect(self._on_source_changed)
        layout.addWidget(self.source_combo)

        # Excel 欄
        self.excel_group = QWidget()
        excel_vl = QVBoxLayout(self.excel_group)
        excel_vl.setContentsMargins(0, 0, 0, 0)
        excel_vl.addWidget(QLabel("Excelファイル"))
        excel_hl = QHBoxLayout()
        self.excel_input = QLineEdit()
        btn_excel = QPushButton("参照")
        btn_excel.clicked.connect(self.browse_excel)
        excel_hl.addWidget(self.excel_input)
        excel_hl.addWidget(btn_excel)
        excel_vl.addLayout(excel_hl)
        self.gen_excel_btn = QPushButton("Excelを生成")
        self.gen_excel_btn.clicked.connect(self.generate_excel)
        excel_vl.addWidget(self.gen_excel_btn)
        layout.addWidget(self.excel_group)

        # Google Sheets 欄
        self.sheets_group = QWidget()
        sheets_vl = QVBoxLayout(self.sheets_group)
        sheets_vl.setContentsMargins(0, 0, 0, 0)
        sheets_vl.addWidget(QLabel("スプレッドシート URL"))
        self.sheets_url_input = QLineEdit()
        self.sheets_url_input.setPlaceholderText("https://docs.google.com/spreadsheets/d/...")
        sheets_vl.addWidget(self.sheets_url_input)

        # credentials.json 配置ボタン＋状態表示
        cred_hl = QHBoxLayout()
        self.cred_btn = QPushButton("📄 credentials.json を配置")
        self.cred_btn.clicked.connect(self._choose_credentials_file)
        self.cred_status_label = QLabel()
        cred_hl.addWidget(self.cred_btn)
        cred_hl.addWidget(self.cred_status_label)
        cred_hl.addStretch()
        sheets_vl.addLayout(cred_hl)

        # Google 認可ボタン＋状態表示
        auth_hl = QHBoxLayout()
        self.auth_btn = QPushButton("🔐 Google で認可する")
        self.auth_btn.clicked.connect(self._authorize_google)
        self.auth_status_label = QLabel()
        auth_hl.addWidget(self.auth_btn)
        auth_hl.addWidget(self.auth_status_label)
        auth_hl.addStretch()
        sheets_vl.addLayout(auth_hl)

        layout.addWidget(self.sheets_group)
        self.sheets_group.setVisible(False)   # 初期は非表示
        self._update_credentials_status()
        self._update_auth_status()

        # ---- Webhook URL ----
        layout.addWidget(QLabel("──────── 通知先 ────────"))

        webhook_label = QLabel(
            "Webhook URL（Discord / Slack / Teams / Chatwork / Google Chat）\n"
            "自動判定されます。画像送信：Discord/Teams/Google Chat 対応"
        )
        webhook_label.setWordWrap(True)
        layout.addWidget(webhook_label)
        self.webhook_input = QLineEdit()
        self.webhook_input.setPlaceholderText("https://discord.com/api/webhooks/...")
        layout.addWidget(self.webhook_input)

        # ---- 締切前日数 ----
        layout.addWidget(QLabel("締切何日前に通知"))
        self.days_spin = QSpinBox()
        self.days_spin.setRange(0, 60)
        layout.addWidget(self.days_spin)

        # ---- メンション ----
        self.mention_checkbox = QCheckBox("メンションを有効（任意）")
        layout.addWidget(self.mention_checkbox)
        mention_box = QWidget()
        mention_l = QVBoxLayout(mention_box)
        mention_l.setContentsMargins(20, 0, 0, 0)
        self.mention_layout = mention_l
        self.mention_layout.addWidget(RowInput("担当名", "ユーザーID", self.mention_layout, False))
        layout.addWidget(mention_box)

        # ---- 確認待ち通知先（レビュアー） ----
        layout.addWidget(QLabel("──────── 確認待ち通知 ────────"))
        self.reviewer_checkbox = QCheckBox("確認待ちタスクを別の人に通知する（任意）")
        layout.addWidget(self.reviewer_checkbox)
        self.reviewer_checkbox.stateChanged.connect(self._on_reviewer_toggled)

        self.reviewer_group = QWidget()
        reviewer_vl = QVBoxLayout(self.reviewer_group)
        reviewer_vl.setContentsMargins(20, 0, 0, 0)

        reviewer_webhook_label = QLabel(
            "レビュアー通知先 Webhook URL\n"
            "（同じプラットフォーム対応。空欄で上の URL を使用）"
        )
        reviewer_webhook_label.setWordWrap(True)
        reviewer_vl.addWidget(reviewer_webhook_label)
        self.reviewer_webhook_input = QLineEdit()
        self.reviewer_webhook_input.setPlaceholderText("省略可（空欄 = 同じ Webhook に送信）")
        reviewer_vl.addWidget(self.reviewer_webhook_input)

        reviewer_vl.addWidget(QLabel("レビュアーのメンション設定（担当名 → レビュアーID）"))
        self.reviewer_mention_layout = QVBoxLayout()
        self.reviewer_mention_layout.addWidget(
            RowInput("担当名", "レビュアーID", self.reviewer_mention_layout, False)
        )
        reviewer_mention_box = QWidget()
        reviewer_mention_box.setLayout(self.reviewer_mention_layout)
        reviewer_vl.addWidget(reviewer_mention_box)

        layout.addWidget(self.reviewer_group)
        self.reviewer_group.setVisible(False)

        # ---- 自動連絡 ----
        layout.addWidget(QLabel("──────── 自動連絡 ────────"))
        self.auto_checkbox = QCheckBox("自動連絡を有効（任意）")
        layout.addWidget(self.auto_checkbox)

        t_l = QHBoxLayout()
        t_l.addWidget(QLabel("連絡時刻"))
        self.time_edit = QTimeEdit(QTime(9, 0))
        t_l.addWidget(self.time_edit)
        layout.addLayout(t_l)

        i_l = QHBoxLayout()
        i_l.addWidget(QLabel("連絡頻度（日）"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 30)
        i_l.addWidget(self.interval_spin)
        layout.addLayout(i_l)

        layout.addWidget(QLabel("制作期間"))
        d_l = QHBoxLayout()
        self.start_date = QDateEdit(QDate.currentDate())
        self.start_date.setCalendarPopup(True)
        self.end_date = QDateEdit(QDate.currentDate().addYears(1))
        self.end_date.setCalendarPopup(True)
        d_l.addWidget(QLabel("開始日"))
        d_l.addWidget(self.start_date)
        d_l.addWidget(QLabel("終了日"))
        d_l.addWidget(self.end_date)
        layout.addLayout(d_l)

        # ---- ボタン ----
        self.save_btn = QPushButton("新規作成")
        self.run_btn  = QPushButton("通知テスト実行")
        self.list_btn = QPushButton("締切教官管理")
        layout.addWidget(self.save_btn)
        layout.addWidget(self.run_btn)
        layout.addWidget(self.list_btn)

        self.save_btn.clicked.connect(self.save_config)
        self.run_btn.clicked.connect(self.run_notify)
        self.list_btn.clicked.connect(self.open_task_list)

        self.config = {}

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

    def _authorize_google(self):
        """Google Sheets の認可フローを別スレッドで実行"""
        def _callback(success, message):
            if success:
                QMessageBox.information(self, "Google Sheets 認可", message)
                self._update_auth_status()
            else:
                QMessageBox.warning(self, "Google Sheets 認可", message)

        t = threading.Thread(
            target=lambda: google_auth_helper.authorize_google_sheets(_callback),
            daemon=True
        )
        t.start()

    def _update_auth_status(self):
        """認可状態を表示"""
        if google_auth_helper.has_token():
            self.auth_status_label.setText("✅ 認可済み")
            self.auth_btn.setEnabled(False)
        else:
            self.auth_status_label.setText("❌ 未認可")
            self.auth_btn.setEnabled(True)

    # ---- データソース切り替え ----
    def _on_source_changed(self, index):
        is_sheets = (index == 1)
        self.excel_group.setVisible(not is_sheets)
        self.sheets_group.setVisible(is_sheets)

    # ---- レビュアー欄 表示切り替え ----
    def _on_reviewer_toggled(self, state):
        self.reviewer_group.setVisible(state == Qt.CheckState.Checked.value)

    def browse_excel(self):
        p, _ = QFileDialog.getOpenFileName(self, "Excelを選択", "", "Excel (*.xlsx *.xls)")
        if p:
            self.excel_input.setText(p)

    def open_manual(self):
        try:
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            pdf_path = os.path.join(base_dir, "SimekiriKyokan_Manual.pdf")
            if os.path.exists(pdf_path):
                os.startfile(pdf_path)
            else:
                QMessageBox.warning(self, "エラー", "マニュアルPDFが見つかりません")
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"マニュアルを開けませんでした:\n{e}")

    def generate_excel(self):
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "Tasks.xlsx")
        if not os.path.exists(template_path):
            QMessageBox.critical(self, "エラー", f"テンプレート Excel が見つかりません:\n{template_path}")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "Excelを保存する場所を選択", "Tasks.xlsx", "Excel Files (*.xlsx)")
        if not save_path:
            return
        if os.path.exists(save_path):
            reply = QMessageBox.question(
                self, "上書き確認",
                f"既存のファイルが存在します。\n上書きしますか？\n\n{save_path}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        try:
            shutil.copyfile(template_path, save_path)
            QMessageBox.information(self, "完了", f"Excelを生成しました:\n{save_path}")
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
        title = self.title_input.text()
        if not title:
            QMessageBox.warning(self, "入力エラー", "締切名を入力してください")
            return

        is_sheets = (self.source_combo.currentIndex() == 1)

        if is_sheets:
            if not self.sheets_url_input.text():
                QMessageBox.warning(self, "入力エラー", "スプレッドシート URL を入力してください")
                return
        else:
            if not self.excel_input.text():
                QMessageBox.warning(self, "入力エラー", "Excelファイルを指定してください")
                return

        if not self.webhook_input.text():
            QMessageBox.warning(self, "入力エラー", "Webhook URLを入力してください")
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
                    QMessageBox.information(self, "保存完了", "設定を保存し、タスクを更新しました（管理者権限あり）")
                else:
                    reply = QMessageBox.question(
                        self, "管理者権限確認",
                        "タスク登録には管理者権限が必要です。昇格しますか？",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.Yes:
                        relaunch_as_admin(config_path, ADMIN_FLAG)
                        QMessageBox.information(self, "保存完了", "設定を保存しました。管理者権限でタスク登録が行われます。")
                    else:
                        QMessageBox.information(self, "保存完了", "設定を保存しました（タスク登録は未実行）")
            except Exception as e:
                QMessageBox.warning(self, "エラー", f"タスクの更新に失敗しました: {e}")
        else:
            QMessageBox.information(self, "保存完了", "設定を保存しました")

        self._reset_form()

    def _reset_form(self):
        self.config = {}
        self.title_input.clear()
        self.excel_input.clear()
        self.sheets_url_input.clear()
        self.webhook_input.clear()
        self.days_spin.setValue(3)
        self.mention_checkbox.setChecked(False)

        for layout in (self.mention_layout, self.reviewer_mention_layout):
            while layout.count():
                w = layout.takeAt(0).widget()
                if w:
                    w.setParent(None)

        self.mention_layout.addWidget(RowInput("担当名", "ユーザーID", self.mention_layout, False))
        self.reviewer_mention_layout.addWidget(RowInput("担当名", "レビュアーID", self.reviewer_mention_layout, False))

        self.reviewer_checkbox.setChecked(False)
        self.reviewer_webhook_input.clear()
        self.auto_checkbox.setChecked(False)
        self.time_edit.setTime(QTime(9, 0))
        self.interval_spin.setValue(1)
        self.source_combo.setCurrentIndex(0)

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

    def run_as_admin_and_register(self):
        if not is_admin():
            relaunch_as_admin(get_config_path(self.config["deadline_id"]), ADMIN_FLAG)
            return
        register_task_admin(self.config)

    def check_first_run_task(self):
        if "deadline_id" not in self.config:
            return
        if not self.config.get("auto_notify"):
            return
        if task_exists(self.config["deadline_id"]):
            return
        reply = QMessageBox.question(
            self, "自動起動の設定", "タスクスケジューラに登録しますか？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.run_as_admin_and_register()

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
