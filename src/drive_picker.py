#drive_picker.py
#
# Google Drive 上のスプレッドシート／フォルダを一覧から選ぶダイアログ。
# URL を手で貼り付けなくても選べるようにするためのもの。

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QFrame,
)

import google_drive
from theme import apply_theme


class DrivePickerDialog(QDialog):
    """
    mode="spreadsheet" … スプレッドシートを選ぶ
    mode="folder"      … フォルダを選ぶ

    選択結果は self.selected（{"id","name","url"}）に入る。
    一覧の取得は時間がかかるのでバックグラウンドで行い、
    結果はシグナル経由で GUI スレッドに渡す。
    """

    items_loaded = pyqtSignal(list)
    load_failed = pyqtSignal(str)

    def __init__(self, mode: str = "spreadsheet", parent=None):
        super().__init__(parent)
        self.mode = mode
        self.selected = None
        self._items = []

        is_folder = (mode == "folder")
        self.setWindowTitle("Google Drive から選択" + ("（フォルダ）" if is_folder else "（スプレッドシート）"))
        self.resize(620, 460)
        apply_theme(self)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(10)

        title = QLabel("フォルダを選択" if is_folder else "スプレッドシートを選択")
        title.setStyleSheet("font-size:15px; font-weight:700;")
        layout.addWidget(title)

        search_row = QHBoxLayout(); search_row.setSpacing(6)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("名前で絞り込み（空欄ならすべて）")
        self.search_input.returnPressed.connect(self.reload)
        self.search_btn = QPushButton("🔍 検索")
        self.search_btn.setFixedWidth(90)
        self.search_btn.clicked.connect(self.reload)
        search_row.addWidget(self.search_input)
        search_row.addWidget(self.search_btn)
        layout.addLayout(search_row)

        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["名前", "更新日時"])
        self.table.setColumnWidth(0, 380)
        self.table.setColumnWidth(1, 160)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.doubleClicked.connect(self._accept_selection)
        layout.addWidget(self.table)

        self.status_label = QLabel()
        self.status_label.setObjectName("desc_lbl")
        layout.addWidget(self.status_label)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)

        btn_row = QHBoxLayout(); btn_row.setSpacing(8)
        btn_row.addStretch()
        cancel_btn = QPushButton("キャンセル")
        cancel_btn.clicked.connect(self.reject)
        self.ok_btn = QPushButton("これを使う")
        self.ok_btn.setObjectName("btn_primary")
        self.ok_btn.clicked.connect(self._accept_selection)
        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(self.ok_btn)
        layout.addLayout(btn_row)

        self.items_loaded.connect(self._on_items_loaded)
        self.load_failed.connect(self._on_load_failed)
        self.reload()

    # ---- 一覧の取得 ----
    def reload(self):
        self.status_label.setText("Google Drive から取得中…")
        self.table.setRowCount(0)
        self.search_btn.setEnabled(False)
        self.ok_btn.setEnabled(False)

        name_filter = self.search_input.text()
        lister = google_drive.list_folders if self.mode == "folder" else google_drive.list_spreadsheets

        import threading

        def _load():
            try:
                items = lister(name_filter)
            except google_drive.DriveError as e:
                self.load_failed.emit(str(e))
                return
            except Exception as e:
                self.load_failed.emit(f"取得に失敗しました:\n{e}")
                return
            self.items_loaded.emit(items)

        threading.Thread(target=_load, daemon=True).start()

    def _on_items_loaded(self, items: list):
        self._items = items
        self.search_btn.setEnabled(True)
        self.ok_btn.setEnabled(True)
        self.table.setRowCount(0)
        for row, item in enumerate(items):
            self.table.insertRow(row)
            for col, value in enumerate([item.get("name", ""), item.get("modified", "")]):
                cell = QTableWidgetItem(value)
                cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, col, cell)
        if items:
            self.table.selectRow(0)
            self.status_label.setText(f"{len(items)} 件見つかりました")
        else:
            self.status_label.setText("該当するものが見つかりませんでした")

    def _on_load_failed(self, message: str):
        self.search_btn.setEnabled(True)
        self.ok_btn.setEnabled(True)
        self.status_label.setText("取得に失敗しました")
        QMessageBox.warning(self, "Google Drive", message)

    # ---- 選択 ----
    def _accept_selection(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self._items):
            QMessageBox.information(self, "選択してください", "一覧から選んでください。")
            return
        self.selected = self._items[row]
        self.accept()


def pick_spreadsheet(parent=None):
    """スプレッドシートを選ばせ、選ばれたら {"id","name","url"} を返す。"""
    dlg = DrivePickerDialog("spreadsheet", parent)
    if dlg.exec() == QDialog.DialogCode.Accepted:
        return dlg.selected
    return None


def pick_folder(parent=None):
    """フォルダを選ばせ、選ばれたら {"id","name","url"} を返す。"""
    dlg = DrivePickerDialog("folder", parent)
    if dlg.exec() == QDialog.DialogCode.Accepted:
        return dlg.selected
    return None
