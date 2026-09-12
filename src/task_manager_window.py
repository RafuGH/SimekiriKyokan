#task_manager_window.py
#
# タスク管理ウィンドウ（登録済み締切教官タスクの一覧・編集・削除・即時実行）

import json
import os
import threading
from datetime import datetime
from functools import partial

from PyQt6.QtCore import QMetaObject, Qt, pyqtSlot
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem,
    QPushButton, QMessageBox, QDialog,
)

import simekiri_notify
from theme import apply_theme
from task_scheduler import (
    TASK_BASE_NAME, get_config_path, get_simekiri_tasks, is_admin, relaunch_as_admin,
)
from task_edit_dialog import TaskEditDialog


class TaskManagerWindow(QWidget):
    def __init__(self, main_app):
        super().__init__()
        self.main_app = main_app
        self.setWindowTitle("締切教官 – 管理")
        self.resize(900, 460)

        apply_theme(self)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 20)
        layout.setSpacing(12)

        hdr_row = QHBoxLayout()
        hdr = QLabel("登録済みの教官一覧")
        hdr.setStyleSheet("font-size:15px; font-weight:700;")
        self.refresh_btn = QPushButton("🔄  更新")
        self.refresh_btn.setFixedWidth(90)
        self.refresh_btn.clicked.connect(self.load_tasks)
        hdr_row.addWidget(hdr)
        hdr_row.addStretch()
        hdr_row.addWidget(self.refresh_btn)
        layout.addLayout(hdr_row)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["教官名", "状態", "次回実行", "最終実行", "結果", "操作"])
        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 70)
        self.table.setColumnWidth(2, 150)
        self.table.setColumnWidth(3, 150)
        self.table.setColumnWidth(4, 70)
        self.table.setColumnWidth(5, 260)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)
        self.load_tasks()

    @staticmethod
    def _format_task_time(raw: str) -> str:
        """
        タスクスケジューラの日時文字列を整形する。
        未実行時に返される 1999-11-30 はWindows の未実行デフォルト値なので
        「未実行」と表示する。
        """
        if not raw or raw.strip() == "":
            return "－"
        # 1999-11-30 はWindowsタスクスケジューラの「未実行」デフォルト値
        if raw.startswith("1999-11-30"):
            return "未実行"
        try:
            # "2025-06-01 09:00:00+09:00" → "2025-06-01 09:00"
            dt_str = raw.split("+")[0].split(".")[0].strip()
            return datetime.fromisoformat(dt_str).strftime("%Y-%m-%d %H:%M")
        except Exception:
            return raw

    def load_tasks(self):
        self.table.setRowCount(0)
        for row, t in enumerate(get_simekiri_tasks()):
            display_name = t["name"]
            if t["name"].startswith(TASK_BASE_NAME + "_"):
                deadline_id_part = t["name"][len(TASK_BASE_NAME)+1:]
                cfg_path = get_config_path(deadline_id_part)
                if os.path.exists(cfg_path):
                    try:
                        with open(cfg_path, "r", encoding="utf-8") as f:
                            display_name = json.load(f).get("title", display_name)
                    except Exception:
                        pass

            self.table.insertRow(row)
            if not t["enabled"]:
                status_text = "🔴 無効"
            elif t["state"] == 3:
                status_text = "🟢 有効"
            else:
                status_text = "⚪ 実行中"

            for col, val in enumerate([
                display_name, status_text,
                self._format_task_time(t["next_run"]),
                self._format_task_time(t["last_run"]),
                str(t["last_result"]),
            ]):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, col, item)

            btn_widget = QWidget()
            btn_layout = QHBoxLayout(btn_widget)
            btn_layout.setContentsMargins(6, 3, 6, 3)
            btn_layout.setSpacing(6)
            edit_btn   = QPushButton("編集")
            edit_btn.setFixedWidth(60)
            delete_btn = QPushButton("削除")
            delete_btn.setFixedWidth(60)
            delete_btn.setObjectName("btn_danger")
            run_btn    = QPushButton("▶ 今すぐ実行")
            edit_btn.clicked.connect(partial(self.edit_task, t["name"]))
            delete_btn.clicked.connect(partial(self.delete_task, t["name"]))
            run_btn.clicked.connect(partial(self.run_task, t["name"]))
            btn_layout.addWidget(edit_btn)
            btn_layout.addWidget(delete_btn)
            btn_layout.addWidget(run_btn)
            self.table.setCellWidget(row, 5, btn_widget)
        self.table.resizeRowsToContents()

    def delete_task(self, task_name):
        reply = QMessageBox.question(
            self, "削除確認", f"「{task_name}」を削除しますか？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        deadline_id = task_name.replace(TASK_BASE_NAME + "_", "")
        config_path = get_config_path(deadline_id)
        if not is_admin():
            relaunch_as_admin(config_path, "--delete")
            return
        try:
            import win32com.client
            service = win32com.client.Dispatch("Schedule.Service")
            service.Connect()
            service.GetFolder("\\").DeleteTask(task_name, 0)
            if os.path.exists(config_path):
                os.remove(config_path)
            QMessageBox.information(self, "削除完了", "削除しました")
        except Exception as e:
            QMessageBox.warning(self, "削除失敗", str(e))
        self.load_tasks()

    def run_task(self, task_name):
        """
        タスクスケジューラ経由ではなく、config を直接読んで
        simekiri_notify.run_notify() を呼び出す。
        スケジュール期間外でも即時実行できる。
        """
        deadline_id = task_name.replace(TASK_BASE_NAME + "_", "")
        config_path = get_config_path(deadline_id)

        if not os.path.exists(config_path):
            QMessageBox.warning(self, "実行失敗", "設定ファイルが見つかりません")
            return

        # 別スレッドで実行してGUIをブロックしない。
        # 完了通知はQMetaObject.invokeMethod経由でGUIスレッドに安全にディスパッチする
        # (QTimer.singleShotは呼び出し元スレッドにイベントループが無いと発火しないため使わない)
        def _run():
            try:
                result = simekiri_notify.run_notify(config_path, test_mode=False)
                slot = "_show_run_success" if result == 0 else "_show_run_error"
                QMetaObject.invokeMethod(self, slot, Qt.ConnectionType.QueuedConnection)
            except Exception as e:
                print("run_task error:", e)

        threading.Thread(target=_run, daemon=True).start()
        QMessageBox.information(self, "送信開始", "通知を送信しています…\n完了後にメッセージが表示されます。")

    @pyqtSlot()
    def _show_run_success(self):
        QMessageBox.information(self, "実行完了", "通知を送信しました ✅")

    @pyqtSlot()
    def _show_run_error(self):
        QMessageBox.warning(self, "実行失敗", "通知の送信中にエラーが発生しました。\nログを確認してください。")

    def edit_task(self, task_name):
        deadline_id = task_name.replace(TASK_BASE_NAME + "_", "")
        config_path = get_config_path(deadline_id)
        if not os.path.exists(config_path):
            QMessageBox.warning(self, "エラー", "設定ファイルが見つかりません")
            return
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        dlg = TaskEditDialog(cfg, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            with open(config_path, "r", encoding="utf-8") as f:
                updated_cfg = json.load(f)
            if hasattr(self, "main_app") and self.main_app:
                self.main_app.update_task(updated_cfg)
            QMessageBox.information(self, "保存完了", f"「{updated_cfg.get('title','')}」を更新しました")
            self.load_tasks()
