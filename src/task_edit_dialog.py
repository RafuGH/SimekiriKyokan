#task_edit_dialog.py
#
# 既存タスクの設定編集ダイアログ

import json
import threading

from PyQt6.QtCore import QTime, QDate, Qt
from PyQt6.QtWidgets import (
    QDialog, QScrollArea, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QComboBox, QCheckBox, QSpinBox, QTimeEdit, QDateEdit,
    QPushButton, QFileDialog, QMessageBox,
)

import google_auth_helper
from task_scheduler import get_config_path
from widgets import RowInput


class TaskEditDialog(QDialog):
    def __init__(self, cfg, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.setWindowTitle(f"タスク設定編集: {cfg.get('title','')}")
        self.resize(540, 640)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        inner = QWidget()
        layout = QVBoxLayout(inner)
        scroll_area.setWidget(inner)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll_area)

        # ---- データソース ----
        layout.addWidget(QLabel("──────── データソース ────────"))
        self.source_combo = QComboBox()
        self.source_combo.addItems(["Excel ファイル", "Google スプレッドシート"])
        current_source = cfg.get("data_source", "excel")
        self.source_combo.setCurrentIndex(1 if current_source == "sheets" else 0)
        self.source_combo.currentIndexChanged.connect(self._on_source_changed)
        layout.addWidget(self.source_combo)

        # Excel 欄
        self.excel_group = QWidget()
        excel_vl = QVBoxLayout(self.excel_group)
        excel_vl.setContentsMargins(0, 0, 0, 0)
        excel_vl.addWidget(QLabel("Excelファイル"))
        excel_hl = QHBoxLayout()
        self.excel_input = QLineEdit(cfg.get("excel_path", ""))
        browse_btn = QPushButton("参照")
        browse_btn.clicked.connect(self.browse_excel)
        excel_hl.addWidget(self.excel_input)
        excel_hl.addWidget(browse_btn)
        excel_vl.addLayout(excel_hl)
        layout.addWidget(self.excel_group)

        # Sheets 欄
        self.sheets_group = QWidget()
        sheets_vl = QVBoxLayout(self.sheets_group)
        sheets_vl.setContentsMargins(0, 0, 0, 0)
        sheets_vl.addWidget(QLabel("スプレッドシート URL"))
        self.sheets_url_input = QLineEdit(cfg.get("sheets_url", ""))
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
        self._on_source_changed(self.source_combo.currentIndex())
        self._update_credentials_status()
        self._update_auth_status()

        # ---- Webhook ----
        layout.addWidget(QLabel("──────── 通知先 ────────"))
        layout.addWidget(QLabel(
            "Webhook URL（Discord / Slack / Teams / Chatwork / Google Chat）\n"
            "画像送信対応：Discord・Teams・Google Chat"
        ))
        self.webhook_input = QLineEdit(cfg.get("webhook_url", ""))
        layout.addWidget(self.webhook_input)

        self.days_spin = QSpinBox()
        self.days_spin.setRange(0, 60)
        self.days_spin.setValue(cfg.get("days_before_deadline", 3))
        layout.addWidget(QLabel("締切何日前に通知"))
        layout.addWidget(self.days_spin)

        # ---- メンション ----
        self.mention_checkbox = QCheckBox("メンションを有効（任意）")
        self.mention_checkbox.setChecked(cfg.get("mention_enabled", False))
        layout.addWidget(self.mention_checkbox)

        self.mention_box = QWidget()
        self.mention_layout = QVBoxLayout(self.mention_box)
        self.mention_layout.setContentsMargins(0, 0, 0, 0)
        mentions = cfg.get("mentions", [])
        if mentions:
            for m in mentions:
                row = RowInput("担当名", "ユーザーID", self.mention_layout)
                row.short.setText(m.get("name", ""))
                row.long.setText(m.get("id", ""))
                self.mention_layout.addWidget(row)
        else:
            self.mention_layout.addWidget(RowInput("担当名", "ユーザーID", self.mention_layout, deletable=False))
        layout.addWidget(self.mention_box)
        for i in range(self.mention_layout.count()):
            w = self.mention_layout.itemAt(i).widget()
            if isinstance(w, RowInput):
                w.update_delete_state()

        # ---- 確認待ち通知先 ----
        layout.addWidget(QLabel("──────── 確認待ち通知 ────────"))
        self.reviewer_checkbox = QCheckBox("確認待ちタスクを別の人に通知する（任意）")
        self.reviewer_checkbox.setChecked(cfg.get("reviewer_enabled", False))
        self.reviewer_checkbox.stateChanged.connect(self._on_reviewer_toggled)
        layout.addWidget(self.reviewer_checkbox)

        self.reviewer_group = QWidget()
        reviewer_vl = QVBoxLayout(self.reviewer_group)
        reviewer_vl.setContentsMargins(20, 0, 0, 0)
        reviewer_vl.addWidget(QLabel(
            "レビュアー通知先 Webhook URL（省略可）\n"
            "Discord/Slack/Teams/Chatwork/Google Chat 対応"
        ))
        self.reviewer_webhook_input = QLineEdit(cfg.get("reviewer_webhook_url", ""))
        reviewer_vl.addWidget(self.reviewer_webhook_input)
        reviewer_vl.addWidget(QLabel("レビュアーのメンション設定（担当名 → レビュアーID）"))
        self.reviewer_mention_layout = QVBoxLayout()
        reviewer_mentions = cfg.get("reviewer_mentions", [])
        if reviewer_mentions:
            for m in reviewer_mentions:
                row = RowInput("担当名", "レビュアーID", self.reviewer_mention_layout)
                row.short.setText(m.get("name", ""))
                row.long.setText(m.get("id", ""))
                self.reviewer_mention_layout.addWidget(row)
        else:
            self.reviewer_mention_layout.addWidget(
                RowInput("担当名", "レビュアーID", self.reviewer_mention_layout, deletable=False)
            )
        reviewer_mention_box = QWidget()
        reviewer_mention_box.setLayout(self.reviewer_mention_layout)
        reviewer_vl.addWidget(reviewer_mention_box)
        layout.addWidget(self.reviewer_group)
        self.reviewer_group.setVisible(cfg.get("reviewer_enabled", False))

        # ---- 自動通知 ----
        layout.addWidget(QLabel("──────── 自動連絡 ────────"))
        self.auto_checkbox = QCheckBox("自動通知を有効（任意）")
        self.auto_checkbox.setChecked(cfg.get("auto_notify", False))
        layout.addWidget(self.auto_checkbox)

        self.time_edit = QTimeEdit(QTime.fromString(cfg.get("notify_time", "09:00"), "HH:mm"))
        t_layout = QHBoxLayout()
        t_layout.addWidget(QLabel("通知時刻"))
        t_layout.addWidget(self.time_edit)
        layout.addLayout(t_layout)

        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 30)
        self.interval_spin.setValue(cfg.get("notify_interval_days", 1))
        i_layout = QHBoxLayout()
        i_layout.addWidget(QLabel("通知頻度（日）"))
        i_layout.addWidget(self.interval_spin)
        layout.addLayout(i_layout)

        # ---- 制作期間 ----
        layout.addWidget(QLabel("制作期間"))
        d_layout = QHBoxLayout()

        start_str = cfg.get("start_date", QDate.currentDate().toString("yyyy-MM-dd"))
        end_str   = cfg.get("end_date",   QDate.currentDate().addYears(1).toString("yyyy-MM-dd"))
        s_date = QDate.fromString(start_str, "yyyy-MM-dd")
        e_date = QDate.fromString(end_str,   "yyyy-MM-dd")
        if not s_date.isValid():
            s_date = QDate.currentDate()
        if not e_date.isValid():
            e_date = QDate.currentDate().addYears(1)

        self.start_date_edit = QDateEdit(s_date)
        self.start_date_edit.setCalendarPopup(True)
        self.end_date_edit   = QDateEdit(e_date)
        self.end_date_edit.setCalendarPopup(True)

        d_layout.addWidget(QLabel("開始日"))
        d_layout.addWidget(self.start_date_edit)
        d_layout.addWidget(QLabel("終了日"))
        d_layout.addWidget(self.end_date_edit)
        layout.addLayout(d_layout)

        save_btn = QPushButton("設定を保存")
        save_btn.clicked.connect(self.save)
        layout.addWidget(save_btn)

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
        """Google Sheets の認可フロー"""
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

    def _on_source_changed(self, index):
        is_sheets = (index == 1)
        self.excel_group.setVisible(not is_sheets)
        self.sheets_group.setVisible(is_sheets)

    def _on_reviewer_toggled(self, state):
        self.reviewer_group.setVisible(state == Qt.CheckState.Checked.value)

    def browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excelを選択", "", "Excel (*.xlsx *.xls)")
        if path:
            self.excel_input.setText(path)

    def _collect_mentions(self, layout_widget) -> list:
        mentions = []
        for i in range(layout_widget.count()):
            w = layout_widget.itemAt(i).widget()
            if isinstance(w, RowInput):
                short, long_ = w.get()
                if short or long_:
                    mentions.append({"name": short, "id": long_})
        return mentions

    def save(self):
        is_sheets = (self.source_combo.currentIndex() == 1)
        if is_sheets and not self.sheets_url_input.text():
            QMessageBox.warning(self, "入力エラー", "スプレッドシート URL を入力してください")
            return
        if not is_sheets and not self.excel_input.text():
            QMessageBox.warning(self, "入力エラー", "Excelファイルを指定してください")
            return

        self.cfg["data_source"]              = "sheets" if is_sheets else "excel"
        self.cfg["excel_path"]               = self.excel_input.text() if not is_sheets else ""
        self.cfg["sheets_url"]               = self.sheets_url_input.text() if is_sheets else ""
        self.cfg["webhook_url"]              = self.webhook_input.text()
        self.cfg["days_before_deadline"]     = self.days_spin.value()
        self.cfg["mention_enabled"]          = self.mention_checkbox.isChecked()
        self.cfg["mentions"]                 = self._collect_mentions(self.mention_layout)
        self.cfg["reviewer_enabled"]         = self.reviewer_checkbox.isChecked()
        self.cfg["reviewer_webhook_url"]     = self.reviewer_webhook_input.text()
        self.cfg["reviewer_mentions"]        = self._collect_mentions(self.reviewer_mention_layout)
        self.cfg["auto_notify"]              = self.auto_checkbox.isChecked()
        self.cfg["notify_time"]              = self.time_edit.time().toString("HH:mm")
        self.cfg["notify_interval_days"]     = self.interval_spin.value()
        self.cfg["start_date"]               = self.start_date_edit.date().toString("yyyy-MM-dd")
        self.cfg["end_date"]                 = self.end_date_edit.date().toString("yyyy-MM-dd")

        config_path = get_config_path(self.cfg["deadline_id"])
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(self.cfg, f, ensure_ascii=False, indent=2)

        self.accept()
