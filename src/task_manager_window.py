#task_manager_window.py
#
# タスク管理ウィンドウ（登録済み締切教官タスクの一覧・編集・削除・即時実行）

import json
import os
import threading
from functools import partial

from PyQt6.QtCore import QMetaObject, Qt, pyqtSlot
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QMessageBox, QDialog,
)

import simekiri_notify
from task_scheduler import (
    TASK_BASE_NAME, get_config_path, get_simekiri_tasks, is_admin, relaunch_as_admin,
)
from task_edit_dialog import TaskEditDialog


class TaskManagerWindow(QWidget):
    def __init__(self, main_app):
        super().__init__()
        self.main_app = main_app
        self.setWindowTitle("タスク管理")
        self.resize(800, 400)

        layout = QVBoxLayout(self)
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["タスク名", "状態", "次回実行", "最終実行", "結果", "操作"])
        self.table.setColumnWidth(0, 100)
        self.table.setColumnWidth(1, 50)
        self.table.setColumnWidth(2, 160)
        self.table.setColumnWidth(3, 160)
        self.table.setColumnWidth(4, 50)
        self.table.setColumnWidth(5, 220)
        layout.addWidget(self.table)

        self.refresh_btn = QPushButton("更新")
        layout.addWidget(self.refresh_btn)
        self.refresh_btn.clicked.connect(self.load_tasks)
        self.load_tasks()

    @staticmethod
    def _format_task_time(raw: str) -> str:
        """
        タスクスケジューラの日時文字列を整形する。
        未実行時に返される 1999-11-30 はWindows の未実行デフォルト値なので
        「未実行」と表示する。次回実行が過去日時の場合も考慮。
        """
        if not raw or raw.strip() == "":
            return "－"
        # 1999-11-30 はWindowsタスクスケジューラの「未実行」デフォルト値
        if raw.startswith("1999-11-30"):
            return "未実行"
        try:
            # "2025-06-01 09:00:00+09:00" → "2025-06-01 09:00"
            dt_str = raw.split("+")[0].split(".")[0].strip()
            from datetime import datetime as _dt
            dt = _dt.fromisoformat(dt_str)
            return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            return raw

    def load_tasks(self):
        tasks = get_simekiri_tasks()
        self.table.setRowCount(0)
        for row, t in enumerate(tasks):
            display_name = t["name"]
            if t["name"].startswith(TASK_BASE_NAME + "_"):
                deadline_id_part = t["name"][len(TASK_BASE_NAME)+1:]
                cfg_path = get_config_path(deadline_id_part)
                if os.path.exists(cfg_path):
                    try:
                        with open(cfg_path, "r", encoding="utf-8") as f:
                            cfg = json.load(f)
                            display_name = cfg.get("title", display_name)
                    except Exception:
                        pass

            self.table.insertRow(row)
            if not t["enabled"]:
                status_text = "🔴無効"
            elif t["state"] == 3:
                status_text = "🟢有効"
            else:
                status_text = "⚪実行"

            self.table.setItem(row, 0, QTableWidgetItem(display_name))
            self.table.setItem(row, 1, QTableWidgetItem(status_text))
            self.table.setItem(row, 2, QTableWidgetItem(self._format_task_time(t["next_run"])))
            self.table.setItem(row, 3, QTableWidgetItem(self._format_task_time(t["last_run"])))
            self.table.setItem(row, 4, QTableWidgetItem(str(t["last_result"])))

            btn_widget = QWidget()
            btn_layout = QHBoxLayout(btn_widget)
            btn_layout.setContentsMargins(0, 0, 0, 0)
            edit_btn   = QPushButton("編集")
            delete_btn = QPushButton("削除")
            run_btn    = QPushButton("今すぐ実行")
            edit_btn.clicked.connect(partial(self.edit_task, t["name"]))
            delete_btn.clicked.connect(partial(self.delete_task, t["name"]))
            run_btn.clicked.connect(partial(self.run_task, t["name"]))
            btn_layout.addWidget(edit_btn)
            btn_layout.addWidget(delete_btn)
            btn_layout.addWidget(run_btn)
            self.table.setCellWidget(row, 5, btn_widget)

    def delete_task(self, task_name):
        reply = QMessageBox.question(
            self, "教官を削除", f"{task_name} を削除しますか？",
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
            root = service.GetFolder("\\")
            root.DeleteTask(task_name, 0)
            if os.path.exists(config_path):
                os.remove(config_path)
            QMessageBox.information(self, "削除完了", f"{task_name} を削除しました")
        except Exception as e:
            QMessageBox.warning(self, "削除失敗", f"削除に失敗しました:\n{e}")
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

        # 別スレッドで実行してGUIをブロックしない
        def _run():
            try:
                result = simekiri_notify.run_notify(config_path, test_mode=False)
                # GUIスレッドへの通知はシグナルが理想だが、
                # シンプルにメッセージボックスをメインスレッドから呼ぶ
                if result == 0:
                    QMetaObject.invokeMethod(
                        self, "_show_run_success",
                        Qt.ConnectionType.QueuedConnection
                    )
                else:
                    QMetaObject.invokeMethod(
                        self, "_show_run_error",
                        Qt.ConnectionType.QueuedConnection
                    )
            except Exception as e:
                print("run_task error:", e)

        t = threading.Thread(target=_run, daemon=True)
        t.start()
        QMessageBox.information(self, "実行開始", "通知を送信しています…\n完了後にメッセージが表示されます。")

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
            QMessageBox.information(self, "保存完了", f"{updated_cfg.get('title','')} を更新しました")
            self.load_tasks()
