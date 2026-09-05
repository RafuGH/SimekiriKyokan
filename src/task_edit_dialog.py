#task_edit_dialog.py
#
# 既存タスクの設定編集ダイアログ

import json
import os

from PyQt6.QtCore import QTime, QDate, Qt
from PyQt6.QtWidgets import (
    QDialog, QScrollArea, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QComboBox, QCheckBox, QSpinBox, QTimeEdit, QDateEdit,
    QPushButton, QFileDialog, QMessageBox, QFrame,
)

import google_auth_helper
import google_drive
import drive_picker
from theme import apply_theme
from help_widgets import field_row, section_header
from google_auth_mixin import GoogleAuthMixin
from task_scheduler import get_config_path
from widgets import RowInput


class TaskEditDialog(GoogleAuthMixin, QDialog):
    def __init__(self, cfg, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.setWindowTitle(f"編集 – {cfg.get('title','')}")
        self.resize(560, 720)

        apply_theme(self)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(20, 16, 20, 24)
        layout.setSpacing(2)
        scroll.setWidget(inner)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

        # ---- データソース ----
        layout.addWidget(section_header("データソース"))
        layout.addLayout(field_row("読み込み元", "datasource"))
        self.source_combo = QComboBox()
        self.source_combo.addItems(["Excel ファイル", "Google スプレッドシート"])
        self.source_combo.setCurrentIndex(1 if cfg.get("data_source") == "sheets" else 0)
        self.source_combo.currentIndexChanged.connect(self._on_source_changed)
        layout.addWidget(self.source_combo)

        # Excel 欄
        self.excel_group = QWidget()
        evl = QVBoxLayout(self.excel_group)
        evl.setContentsMargins(0, 4, 0, 0)
        evl.setSpacing(4)
        evl.addLayout(field_row("Excel ファイルのパス"))
        ehl = QHBoxLayout(); ehl.setSpacing(6)
        self.excel_input = QLineEdit(cfg.get("excel_path", ""))
        browse_btn = QPushButton("参照")
        browse_btn.setFixedWidth(64)
        browse_btn.clicked.connect(self.browse_excel)
        ehl.addWidget(self.excel_input)
        ehl.addWidget(browse_btn)
        evl.addLayout(ehl)
        layout.addWidget(self.excel_group)

        # Sheets 欄
        self.sheets_group = QWidget()
        svl = QVBoxLayout(self.sheets_group)
        svl.setContentsMargins(0, 4, 0, 0)
        svl.setSpacing(4)
        svl.addLayout(field_row("スプレッドシート URL", "google_auth"))
        sheet_hl = QHBoxLayout(); sheet_hl.setSpacing(6)
        self.sheets_url_input = QLineEdit(cfg.get("sheets_url", ""))
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
        self._init_google_auth()
        self._on_source_changed(self.source_combo.currentIndex())
        self._update_credentials_status()
        self._refresh_auth_label()

        # ---- Webhook ----
        layout.addWidget(section_header("通知先"))
        layout.addLayout(field_row("Webhook URL", "webhook"))
        self.webhook_input = QLineEdit(cfg.get("webhook_url", ""))
        layout.addWidget(self.webhook_input)

        layout.addLayout(field_row("締切何日前に通知", "days_before"))
        self.days_spin = QSpinBox()
        self.days_spin.setRange(0, 60)
        self.days_spin.setValue(cfg.get("days_before_deadline", 3))
        self.days_spin.setFixedWidth(100)
        layout.addWidget(self.days_spin)

        # ---- メンション ----
        layout.addLayout(field_row("担当者メンション", "mention"))
        self.mention_checkbox = QCheckBox("メンションを有効にする")
        self.mention_checkbox.setChecked(cfg.get("mention_enabled", False))
        layout.addWidget(self.mention_checkbox)

        self.mention_box = QWidget()
        self.mention_layout = QVBoxLayout(self.mention_box)
        self.mention_layout.setContentsMargins(20, 0, 0, 0)
        self.mention_layout.setSpacing(2)
        mentions = cfg.get("mentions") or []
        for m in (mentions or [{"name": "", "id": ""}]):
            row = RowInput("担当名", "ユーザーID", self.mention_layout, deletable=bool(mentions))
            row.short.setText(m.get("name", ""))
            row.long.setText(m.get("id", ""))
            self.mention_layout.addWidget(row)
        for i in range(self.mention_layout.count()):
            w = self.mention_layout.itemAt(i).widget()
            if isinstance(w, RowInput):
                w.update_delete_state()
        layout.addWidget(self.mention_box)

        # ---- 確認待ち通知先 ----
        layout.addWidget(section_header("確認待ち通知"))
        layout.addLayout(field_row("レビュアーへの通知", "reviewer"))
        self.reviewer_checkbox = QCheckBox("確認待ちタスクを別の担当者に通知する")
        self.reviewer_checkbox.setChecked(cfg.get("reviewer_enabled", False))
        self.reviewer_checkbox.stateChanged.connect(self._on_reviewer_toggled)
        layout.addWidget(self.reviewer_checkbox)

        self.reviewer_group = QWidget()
        rvl = QVBoxLayout(self.reviewer_group)
        rvl.setContentsMargins(20, 4, 0, 0)
        rvl.setSpacing(4)
        rvl.addLayout(field_row("レビュアー Webhook URL（省略可）"))
        self.reviewer_webhook_input = QLineEdit(cfg.get("reviewer_webhook_url", ""))
        rvl.addWidget(self.reviewer_webhook_input)
        rvl.addLayout(field_row("レビュアーのメンション"))
        self.reviewer_mention_layout = QVBoxLayout()
        self.reviewer_mention_layout.setSpacing(2)
        reviewer_mentions = cfg.get("reviewer_mentions") or []
        for m in (reviewer_mentions or [{"name": "", "id": ""}]):
            row = RowInput("担当名", "レビュアーID", self.reviewer_mention_layout, deletable=bool(reviewer_mentions))
            row.short.setText(m.get("name", ""))
            row.long.setText(m.get("id", ""))
            self.reviewer_mention_layout.addWidget(row)
        reviewer_mention_box = QWidget()
        reviewer_mention_box.setLayout(self.reviewer_mention_layout)
        rvl.addWidget(reviewer_mention_box)
        layout.addWidget(self.reviewer_group)
        self.reviewer_group.setVisible(cfg.get("reviewer_enabled", False))

        # ---- 提出フォルダ監視 ----
        layout.addWidget(section_header("提出フォルダ監視"))
        layout.addLayout(field_row("フォルダへの提出を通知", "submission"))
        self.submission_checkbox = QCheckBox("提出フォルダを監視して通知する")
        self.submission_checkbox.setChecked(cfg.get("submission_watch_enabled", False))
        self.submission_checkbox.stateChanged.connect(self._on_submission_toggled)
        layout.addWidget(self.submission_checkbox)

        self.submission_group = QWidget()
        subvl = QVBoxLayout(self.submission_group)
        subvl.setContentsMargins(20, 4, 0, 0); subvl.setSpacing(4)

        subvl.addLayout(field_row("監視する場所"))
        self.submission_source_combo = QComboBox()
        self.submission_source_combo.addItems(["ローカル / OneDrive の同期フォルダ", "Google Drive のフォルダ"])
        self.submission_source_combo.setCurrentIndex(1 if cfg.get("submission_source") == "drive" else 0)
        self.submission_source_combo.currentIndexChanged.connect(self._on_submission_source_changed)
        subvl.addWidget(self.submission_source_combo)

        # ローカルフォルダ
        self.submission_local_group = QWidget()
        local_vl = QVBoxLayout(self.submission_local_group)
        local_vl.setContentsMargins(0, 0, 0, 0); local_vl.setSpacing(4)
        local_vl.addLayout(field_row("監視するフォルダ"))
        sub_hl = QHBoxLayout(); sub_hl.setSpacing(6)
        self.submission_folder_input = QLineEdit(cfg.get("submission_folder", ""))
        self.submission_folder_input.setPlaceholderText("例：C:\\Users\\名前\\OneDrive - 会社名\\チーム\\提出")
        btn_sub = QPushButton("参照"); btn_sub.setFixedWidth(64)
        btn_sub.clicked.connect(self.browse_submission_folder)
        sub_hl.addWidget(self.submission_folder_input); sub_hl.addWidget(btn_sub)
        local_vl.addLayout(sub_hl)
        self.submission_recursive_checkbox = QCheckBox("サブフォルダも対象にする")
        self.submission_recursive_checkbox.setChecked(cfg.get("submission_recursive", True))
        local_vl.addWidget(self.submission_recursive_checkbox)
        subvl.addWidget(self.submission_local_group)

        # Google Drive フォルダ
        self.submission_drive_group = QWidget()
        drive_vl = QVBoxLayout(self.submission_drive_group)
        drive_vl.setContentsMargins(0, 0, 0, 0); drive_vl.setSpacing(4)
        drive_vl.addLayout(field_row("Google Drive のフォルダ"))
        drive_hl = QHBoxLayout(); drive_hl.setSpacing(6)
        drive_id = cfg.get("submission_drive_folder_id", "")
        self.submission_drive_input = QLineEdit(
            google_drive.FOLDER_URL.format(id=drive_id) if drive_id else ""
        )
        self.submission_drive_input.setPlaceholderText("フォルダURL、または「Drive から選択」")
        self.pick_drive_folder_btn = QPushButton("📂 Drive から選択")
        self.pick_drive_folder_btn.setFixedWidth(150)
        self.pick_drive_folder_btn.clicked.connect(self.pick_submission_folder_from_drive)
        drive_hl.addWidget(self.submission_drive_input)
        drive_hl.addWidget(self.pick_drive_folder_btn)
        drive_vl.addLayout(drive_hl)
        subvl.addWidget(self.submission_drive_group)
        self._submission_drive_folder_name = cfg.get("submission_drive_folder_name", "")
        self._on_submission_source_changed(self.submission_source_combo.currentIndex())
        subvl.addLayout(field_row("対象の拡張子（省略可・カンマ区切り）"))
        self.submission_ext_input = QLineEdit(cfg.get("submission_extensions", ""))
        self.submission_ext_input.setPlaceholderText("例：.xlsx,.docx,.pdf（空欄ならすべて）")
        subvl.addWidget(self.submission_ext_input)
        subvl.addLayout(field_row("提出通知先 Webhook URL（省略可）"))
        self.submission_webhook_input = QLineEdit(cfg.get("submission_webhook_url", ""))
        self.submission_webhook_input.setPlaceholderText("省略すると上の URL を使用")
        subvl.addWidget(self.submission_webhook_input)
        layout.addWidget(self.submission_group)
        self.submission_group.setVisible(cfg.get("submission_watch_enabled", False))

        # ---- 自動通知 ----
        layout.addWidget(section_header("自動連絡"))
        layout.addLayout(field_row("タスクスケジューラ連携", "auto_notify"))
        self.auto_checkbox = QCheckBox("自動送信を有効にする")
        self.auto_checkbox.setChecked(cfg.get("auto_notify", False))
        layout.addWidget(self.auto_checkbox)

        tl = QHBoxLayout(); tl.setSpacing(10)
        tl.addWidget(QLabel("送信時刻"))
        self.time_edit = QTimeEdit(QTime.fromString(cfg.get("notify_time", "09:00"), "HH:mm"))
        self.time_edit.setFixedWidth(100)
        tl.addWidget(self.time_edit)
        tl.addSpacing(20)
        tl.addWidget(QLabel("頻度（日おき）"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 30)
        self.interval_spin.setValue(cfg.get("notify_interval_days", 1))
        self.interval_spin.setFixedWidth(80)
        tl.addWidget(self.interval_spin)
        tl.addStretch()
        layout.addLayout(tl)

        layout.addLayout(field_row("制作期間"))
        dl = QHBoxLayout(); dl.setSpacing(8)

        def _parse_date(s, fallback):
            d = QDate.fromString(s, "yyyy-MM-dd")
            return d if d.isValid() else fallback

        self.start_date_edit = QDateEdit(_parse_date(cfg.get("start_date", ""), QDate.currentDate()))
        self.start_date_edit.setCalendarPopup(True)
        self.end_date_edit = QDateEdit(_parse_date(cfg.get("end_date", ""), QDate.currentDate().addYears(1)))
        self.end_date_edit.setCalendarPopup(True)
        dl.addWidget(QLabel("開始")); dl.addWidget(self.start_date_edit)
        dl.addSpacing(8)
        dl.addWidget(QLabel("終了")); dl.addWidget(self.end_date_edit)
        dl.addStretch()
        layout.addLayout(dl)

        layout.addSpacing(16)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)
        layout.addSpacing(10)
        save_btn = QPushButton("変更を保存する")
        save_btn.setObjectName("btn_primary")
        save_btn.clicked.connect(self.save)
        layout.addWidget(save_btn)
        layout.addStretch()

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
            "submission_watch_enabled": self.submission_checkbox.isChecked(),
            "submission_source":        "drive" if self.submission_source_combo.currentIndex() == 1 else "local",
            "submission_folder":        self.submission_folder_input.text().strip(),
            "submission_drive_folder_id":   google_drive.extract_folder_id(self.submission_drive_input.text()),
            "submission_drive_folder_name": self._submission_drive_folder_name,
            "submission_extensions":    self.submission_ext_input.text().strip(),
            "submission_webhook_url":   self.submission_webhook_input.text().strip(),
            "submission_recursive":     self.submission_recursive_checkbox.isChecked(),
            "auto_notify":          self.auto_checkbox.isChecked(),
            "notify_time":          self.time_edit.time().toString("HH:mm"),
            "notify_interval_days": self.interval_spin.value(),
            "start_date":           self.start_date_edit.date().toString("yyyy-MM-dd"),
            "end_date":             self.end_date_edit.date().toString("yyyy-MM-dd"),
        })

        config_path = get_config_path(self.cfg["deadline_id"])
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(self.cfg, f, ensure_ascii=False, indent=2)

        self.accept()
