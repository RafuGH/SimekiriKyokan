#simekiri_gui.py

import ctypes
import sys, os, json, shutil, traceback
from datetime import datetime, timedelta
import subprocess
import re
import simekiri_notify
import uuid

from PyQt6.QtWidgets import *
from PyQt6.QtCore import QTime, QDate, Qt
from PyQt6.QtGui import QColor
from functools import partial

# ===============================
# 管理者権限
# ===============================
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

ShellExecuteW = ctypes.windll.shell32.ShellExecuteW

# ===============================
# パス
# ===============================
APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
os.makedirs(APP_DIR, exist_ok=True)

def get_config_path(deadline_id):
    return os.path.join(APP_DIR, f"{deadline_id}.json")

ERROR_LOG = os.path.join(APP_DIR, "gui_error_log.txt")

TASK_BASE_NAME = "SimekiriKyokan"

def get_task_name(deadline_id):
    return f"{TASK_BASE_NAME}_{deadline_id}"

ADMIN_FLAG = "--admin-register"


def generate_deadline_id(category, end_date_str, title):
    safe_title = re.sub(r'[^a-zA-Z0-9ぁ-んァ-ン一-龯]', '', title)[:10]
    uid = uuid.uuid4().hex[:6]
    return f"{category}_{safe_title}_{uid}"


# ===============================
# タスク存在確認
# ===============================
def task_exists(deadline_id):
    try:
        import win32com.client
        service = win32com.client.Dispatch("Schedule.Service")
        service.Connect()
        service.GetFolder("\\").GetTask(get_task_name(deadline_id))
        return True
    except:
        return False

# ===============================
# タスク一覧取得
# ===============================
def get_simekiri_tasks():
    import win32com.client

    service = win32com.client.Dispatch("Schedule.Service")
    service.Connect()
    root = service.GetFolder("\\")

    tasks = root.GetTasks(0)

    result = []

    for task in tasks:
        if task.Name.startswith(TASK_BASE_NAME):
            result.append({
                "name": task.Name,
                "state": task.State,
                "enabled": task.Enabled,
                "next_run": str(task.NextRunTime),
                "last_run": str(task.LastRunTime),
                "last_result": task.LastTaskResult
            })

    return result



# ===============================
# タスク有効/無効切り替え
# ===============================
def set_task_enabled(task_name, enabled):
    import win32com.client

    service = win32com.client.Dispatch("Schedule.Service")
    service.Connect()
    root = service.GetFolder("\\")

    task = root.GetTask(task_name)
    task.Enabled = enabled


# ===============================
# タスクごとのconfigパス取得
# ===============================
def get_task_config_path(deadline_id):
    return os.path.join(APP_DIR, f"{deadline_id}.json")


# ===============================
# タスク登録（管理者）
# ===============================
def register_task_admin(config):
    import win32com.client

    task_name = get_task_name(config["deadline_id"])

    service = win32com.client.Dispatch("Schedule.Service")
    service.Connect()
    root = service.GetFolder("\\")

    # 必ず削除してから作り直す
    try:
        root.DeleteTask(task_name, 0)
    except:
        pass

    task_def = service.NewTask(0)

    # ===============================
    # 日時計算
    # ===============================
    h, m = map(int, config["notify_time"].split(":"))
    start_date = datetime.strptime(config["start_date"], "%Y-%m-%d")
    end_date = datetime.strptime(config["end_date"], "%Y-%m-%d")

    start = start_date.replace(hour=h, minute=m, second=0)
    end = end_date.replace(hour=23, minute=59, second=59)

    # ===============================
    # トリガー
    # ===============================
    trigger = task_def.Triggers.Create(2)
    trigger.StartBoundary = start.strftime("%Y-%m-%dT%H:%M:%S")
    trigger.EndBoundary = end.strftime("%Y-%m-%dT%H:%M:%S")
    trigger.DaysInterval = max(1, config["notify_interval_days"])
    trigger.Enabled = True

    # ===============================
    # アクション
    # ===============================
    action = task_def.Actions.Create(0)
    
    if getattr(sys, 'frozen', False):
        # exe実行時
        action.Path = sys.executable
        config_path = get_config_path(config["deadline_id"])
        action.Arguments = f'--notify "{config_path}"'
        action.WorkingDirectory = os.path.dirname(sys.executable)
    else:
        # Python実行時
        action.Path = sys.executable
        config_path = get_config_path(config["deadline_id"])
        action.Arguments = f'"{os.path.abspath(__file__)}" --notify "{config_path}"'
        action.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))

    # ===============================
    # ユーザー設定
    # ===============================
    task_def.Principal.LogonType = 3
    task_def.Principal.RunLevel = 0

    settings = task_def.Settings
    settings.Enabled = True
    settings.StartWhenAvailable = True
    settings.ExecutionTimeLimit = "PT0S"

    # CREATE_OR_UPDATE（6）
    root.RegisterTaskDefinition(
        task_name,
        task_def,
        6,
        None,
        None,
        3
    )


    config["task_registered"] = True


# ===============================
# メンション行
# ===============================
class RowInput(QWidget):
    def __init__(self, short_ph, long_ph, parent_layout, deletable=True):
        super().__init__()
        self.parent_layout = parent_layout
        layout = QHBoxLayout(self)

        self.short = QLineEdit()
        self.short.setPlaceholderText(short_ph)
        self.long = QLineEdit()
        self.long.setPlaceholderText(long_ph)

        self.add_btn = QPushButton("＋")
        self.del_btn = QPushButton("－")

        self.add_btn.clicked.connect(self.add)
        self.del_btn.clicked.connect(self.delete)

        if not deletable:
            self.del_btn.setEnabled(False)

        for w in (self.short, self.long, self.add_btn, self.del_btn):
            layout.addWidget(w)

    def update_delete_state(self):
        self.del_btn.setEnabled(self.parent_layout.count() > 1)

    def add(self):
        row = RowInput("担当名", "ユーザーID", self.parent_layout)
        self.parent_layout.addWidget(row)
        self.update_all()

    def delete(self):
        if self.parent_layout.count() <= 1:
            return
    
        self.parent_layout.removeWidget(self)
        self.deleteLater()
        self.update_all()

    def update_all(self):
        for i in range(self.parent_layout.count()):
            w = self.parent_layout.itemAt(i).widget()
            if isinstance(w, RowInput):
                w.update_delete_state()

    def get(self):
        return self.short.text(), self.long.text()


# ===============================
# GUI
# ===============================
class NotifierApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("締切教官 v2.0")
        self.resize(520, 760)

        layout = QVBoxLayout(self)

        # ===== 締切名＋ヘルプ =====
        title_row = QHBoxLayout()

        title_label = QLabel("締切名")
        
        palette = self.palette()
        base_color = palette.color(palette.ColorRole.Window)
        
        # 明度で判定（128より暗ければダーク）
        is_dark = base_color.lightness() < 128

        self.help_btn = QPushButton("？")
        self.help_btn.setFixedSize(24, 24)
        
        if is_dark:
            # ダークモード（今のまま＋少し明るくなる）
            self.help_btn.setStyleSheet("""
                QPushButton {
                    font-size:14px;
                    font-weight:bold;
                    border:1px solid palette(mid);
                    border-radius:4px;
                    padding:0px;
                    margin-top:0px;
                    background-color: palette(button);
                    color: palette(button-text);
                }
                QPushButton:hover {
                    background-color: palette(light);  /* 少し明るく */
                }
            """)
        else:
            # ライトモード（白ベースのグレー）
            self.help_btn.setStyleSheet("""
                QPushButton {
                    font-size:14px;
                    font-weight:bold;
                    border:1px solid #cccccc;
                    border-radius:4px;
                    padding:0px;
                    margin-top:0px;
                    background-color: #f5f5f5;   /* 白寄りグレー */
                    color: black;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;   /* 少し暗く */
                }
            """)
        title_label.setContentsMargins(0, 20, 0, 0)
        self.help_btn.clicked.connect(self.open_manual)
        
        title_row.addWidget(title_label)
        title_row.addStretch()
        title_row.addWidget(self.help_btn)
        
        layout.addLayout(title_row)
        
        # 締切名入力欄
        self.title_input = QLineEdit()
        layout.addWidget(self.title_input)

        layout.addWidget(QLabel("カテゴリ"))
        self.category_combo = QComboBox()
        self.category_combo.addItems(["report", "game", "school", "work", "personal"])
        layout.addWidget(self.category_combo)


        # Excel
        layout.addWidget(QLabel("Excelファイル"))
        excel_l = QHBoxLayout()
        self.excel_input = QLineEdit()
        btn = QPushButton("参照")
        btn.clicked.connect(self.browse_excel)
        excel_l.addWidget(self.excel_input)
        excel_l.addWidget(btn)
        layout.addLayout(excel_l)
        
        # Excel生成ボタン
        self.gen_excel_btn = QPushButton("Excelを生成")
        self.gen_excel_btn.clicked.connect(self.generate_excel)
        layout.addWidget(self.gen_excel_btn)

        # Webhook
        layout.addWidget(QLabel("Discord Webhook URL"))
        self.webhook_input = QLineEdit()
        layout.addWidget(self.webhook_input)

        # 日前
        layout.addWidget(QLabel("締切何日前に通知"))
        self.days_spin = QSpinBox()
        self.days_spin.setRange(0, 60)
        layout.addWidget(self.days_spin)

        # メンション
        self.mention_checkbox = QCheckBox("メンションを有効（任意）")
        layout.addWidget(self.mention_checkbox)

        mention_box = QWidget()
        mention_l = QVBoxLayout(mention_box)
        mention_l.setContentsMargins(20, 0, 0, 0)
        self.mention_layout = mention_l
        self.mention_layout.addWidget(RowInput("担当名", "ユーザーID", self.mention_layout, False))
        layout.addWidget(mention_box)

        # 自動連絡
        layout.addWidget(QLabel("──────── 自動連絡 ────────"))
        self.auto_checkbox = QCheckBox("自動連絡を有効（任意）")
        layout.addWidget(self.auto_checkbox)

        # 時刻
        t_l = QHBoxLayout()
        t_l.addWidget(QLabel("連絡時刻"))
        self.time_edit = QTimeEdit(QTime(9, 0))
        t_l.addWidget(self.time_edit)
        layout.addLayout(t_l)

        # 頻度
        i_l = QHBoxLayout()
        i_l.addWidget(QLabel("連絡頻度（日）"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 30)
        i_l.addWidget(self.interval_spin)
        layout.addLayout(i_l)

        # 制作期間
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

        # ボタン
        self.save_btn = QPushButton("新規作成")
        self.run_btn = QPushButton("通知テスト実行")
        self.list_btn = QPushButton("締切教官管理")
        layout.addWidget(self.save_btn)
        layout.addWidget(self.run_btn)
        layout.addWidget(self.list_btn)

        self.save_btn.clicked.connect(self.save_config)
        self.run_btn.clicked.connect(self.run_notify)
        self.list_btn.clicked.connect(self.open_task_list)

        self.config = {}

    # ===============================
    # 既存ロジック（完全保持）
    # ===============================
    def run_as_admin_and_register(self):

        if not is_admin():
            params = f'"{sys.executable}" {ADMIN_FLAG}'
            script_path = os.path.abspath(__file__)
            
            if getattr(sys, 'frozen', False):
                # exe版
                config_path = get_config_path(self.config["deadline_id"])
                
                ShellExecuteW(
                    None,
                    "runas",
                    sys.executable,
                    f'{ADMIN_FLAG} "{config_path}"',
                    None,
                    1
                )

            else:
                script_path = os.path.abspath(__file__)
                config_path = get_config_path(self.config["deadline_id"])
            
                ShellExecuteW(
                    None,
                    "runas",
                    sys.executable,
                    f'"{script_path}" {ADMIN_FLAG} "{config_path}"',
                    None,
                    1
                )

            return

        register_task_admin(self.config)

    def check_first_run_task(self):
    
        if "deadline_id" not in self.config:
            return
    
        if not self.config.get("auto_notify"):
            return
    
        # すでにタスクが存在しているなら何もしない
        if task_exists(self.config["deadline_id"]):
            return
    
        reply = QMessageBox.question(
            self,
            "自動起動の設定",
            "タスクスケジューラに登録しますか？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
    
        if reply == QMessageBox.StandardButton.Yes:
            self.run_as_admin_and_register()



    def browse_excel(self):
        p, _ = QFileDialog.getOpenFileName(self, "Excelを選択", "", "Excel (*.xlsx *.xls)")
        if p:
            self.excel_input.setText(p)

    # ===============================
    # マニュアルを開く
    # ===============================
    def open_manual(self):
        try:
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))

            pdf_path = os.path.join(base_dir, "SimekiriKyokan_Manual.pdf")

            if os.path.exists(pdf_path):
                os.startfile(pdf_path)  # Windows標準PDFビューアで開く
            else:
                QMessageBox.warning(self, "エラー", "マニュアルPDFが見つかりません")

        except Exception as e:
            QMessageBox.warning(self, "エラー", f"マニュアルを開けませんでした:\n{e}")
            
    # ===============================
    # Excel生成（任意の場所にコピー）
    # ===============================
    def generate_excel(self):
        # exe / Python 版どちらでも data/Tasks.xlsx 参照
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        template_path = os.path.join(base_dir, "data", "Tasks.xlsx")
        
        if not os.path.exists(template_path):
            QMessageBox.critical(self, "エラー", f"テンプレート Excel が見つかりません:\n{template_path}")
            return
    
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Excelを保存する場所を選択",
            "Tasks.xlsx",
            "Excel Files (*.xlsx)"
        )
        if not save_path:
            return  # キャンセル
    
        # 上書き確認
        if os.path.exists(save_path):
            reply = QMessageBox.question(
                self,
                "上書き確認",
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

    def save_config(self):
    
        title = self.title_input.text()

        if not title:
            QMessageBox.warning(self,"入力エラー","締切名を入力してください")
            return
        
        if not self.excel_input.text():
            QMessageBox.warning(self,"入力エラー","Excelファイルを指定してください")
            return
        
        if not self.webhook_input.text():
            QMessageBox.warning(self,"入力エラー","Webhook URLを入力してください")
            return

        category = self.category_combo.currentText()
        end_date = self.end_date.date().toString("yyyy-MM-dd")
    
        deadline_id = generate_deadline_id(category, end_date, title)
        
        cfg = {
            "deadline_id": deadline_id,
            "title": title,
            "category": category,
            "excel_path": self.excel_input.text(),
            "webhook_url": self.webhook_input.text(),
            "days_before_deadline": self.days_spin.value(),
            "mention_enabled": self.mention_checkbox.isChecked(),
            "auto_notify": self.auto_checkbox.isChecked(),
            "notify_time": self.time_edit.time().toString("HH:mm"),
            "notify_interval_days": self.interval_spin.value(),
            "start_date": self.start_date.date().toString("yyyy-MM-dd"),
            "end_date": end_date
        }
        
        # ===== メンション取得 =====
        mentions = []
        for i in range(self.mention_layout.count()):
            w = self.mention_layout.itemAt(i).widget()
            if isinstance(w, RowInput):
                short, long = w.get()
                if short or long:
                    mentions.append({"name": short, "id": long})
        
        cfg["mentions"] = mentions
    
        config_path = get_config_path(deadline_id)
    
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    
        self.config = cfg
    
        # 自動通知が有効なら古いタスクを削除して再登録
        if cfg["auto_notify"]:
            try:
                import win32com.client
                service = win32com.client.Dispatch("Schedule.Service")
                service.Connect()
                folder = service.GetFolder("\\")
                task_name = get_task_name(deadline_id)
    
                # 管理者権限で登録
                if is_admin():
                    register_task_admin(cfg)
                    QMessageBox.information(self, "保存完了", "設定を保存し、タスクを更新しました（管理者権限あり）")
                else:
                    reply = QMessageBox.question(
                        self,
                        "管理者権限確認",
                        "タスク登録には管理者権限が必要です。昇格しますか？",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.Yes:
                        if getattr(sys, 'frozen', False):
                            ShellExecuteW(
                                None,
                                "runas",
                                sys.executable,
                                f'{ADMIN_FLAG} "{config_path}"',
                                None,
                                1
                            )
                        else:
                            script_path = os.path.abspath(__file__)
                            ShellExecuteW(
                                None,
                                "runas",
                                sys.executable,
                                f'"{script_path}" {ADMIN_FLAG} "{config_path}"',
                                None,
                                1
                            )
                        QMessageBox.information(self, "保存完了", "設定を保存しました。管理者権限でタスク登録が行われます。")
                    else:
                        QMessageBox.information(self, "保存完了", "設定を保存しました（タスク登録は未実行）")
            except Exception as e:
                QMessageBox.warning(self, "エラー", f"タスクの更新に失敗しました: {e}")
        else:
            QMessageBox.information(self, "保存完了", "設定を保存しました")
            
        self.config = {}
        
        self.title_input.clear()
        self.excel_input.clear()
        self.webhook_input.clear()
        self.days_spin.setValue(3)
        
        # ★ 追加
        self.mention_checkbox.setChecked(False)
        
        # メンション行を完全リセット
        while self.mention_layout.count():
            w = self.mention_layout.takeAt(0).widget()
            if w:
                w.setParent(None)
        
        # 初期行を1つだけ追加
        self.mention_layout.addWidget(
            RowInput("担当名", "ユーザーID", self.mention_layout, False)
        )
        
        self.auto_checkbox.setChecked(False)
        self.time_edit.setTime(QTime(9,0))
        self.interval_spin.setValue(1)

    def run_notify(self):
        # 保存済みファイル一覧取得
        files = [f for f in os.listdir(APP_DIR) if f.endswith(".json")]
    
        if not files:
            QMessageBox.warning(self, "エラー", "保存された設定がありません")
            return
    
        # 一番新しいファイルを使う
        files.sort(key=lambda x: os.path.getmtime(os.path.join(APP_DIR, x)), reverse=True)
        config_path = os.path.join(APP_DIR, files[0])
    
        simekiri_notify.run_notify(config_path, test_mode=True)

    def open_task_list(self):
        self.task_list_window = TaskManagerWindow(self)
        self.task_list_window.show()
        
        # ===== タスク再登録 =====
    def update_task(self, cfg):
    
        task_name = get_task_name(cfg["deadline_id"])
        config_path = get_config_path(cfg["deadline_id"])
    
        # ===== 管理者ならそのまま登録 =====
        if is_admin():
            try:
                import win32com.client
                service = win32com.client.Dispatch("Schedule.Service")
                service.Connect()
                folder = service.GetFolder("\\")
    
                try:
                    folder.DeleteTask(task_name, 0)
                except:
                    pass
    
                register_task_admin(cfg)
    
            except Exception as e:
                QMessageBox.warning(self, "タスク更新失敗", str(e))
    
            return  # ★超重要
    
        # ===== 管理者でない場合のみ昇格 =====
        if getattr(sys, 'frozen', False):
            # exe版
            ShellExecuteW(
                None,
                "runas",
                sys.executable,
                f'{ADMIN_FLAG} "{config_path}"',
                None,
                1
            )
        else:
            # python実行版
            script_path = os.path.abspath(__file__)
            ShellExecuteW(
                None,
                "runas",
                sys.executable,
                f'"{script_path}" {ADMIN_FLAG} "{config_path}"',
                None,
                1
            )



class TaskManagerWindow(QWidget):
    def __init__(self, main_app):
        super().__init__()   # ← parent渡さない
        self.main_app = main_app
        self.setWindowTitle("タスク管理")
        self.resize(800, 400)

        layout = QVBoxLayout(self)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["タスク名", "状態", "次回実行", "最終実行", "結果", "操作"]
        )

        self.table.setColumnWidth(0, 100)  # タスク名
        self.table.setColumnWidth(1, 50)   # 状態
        self.table.setColumnWidth(2, 160)  # 次回実行
        self.table.setColumnWidth(3, 160)  # 最終実行
        self.table.setColumnWidth(4, 50)   # 結果
        self.table.setColumnWidth(5, 220)  # 操作

        layout.addWidget(self.table)

        self.refresh_btn = QPushButton("更新")
        layout.addWidget(self.refresh_btn)

        self.refresh_btn.clicked.connect(self.load_tasks)

        self.load_tasks()

    def load_tasks(self):
        tasks = get_simekiri_tasks()
        self.table.setRowCount(0)

        for row, t in enumerate(tasks):
            display_name = t["name"]  # とりあえずタスク名をデフォルト表示
    
            # タスク名から deadline_id を復元
            if t["name"].startswith(TASK_BASE_NAME + "_"):
                deadline_id_part = t["name"][len(TASK_BASE_NAME)+1:]
                cfg_path = get_task_config_path(deadline_id_part)
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
            self.table.setItem(row, 0, QTableWidgetItem(display_name))  # ← タスク名列追加
            self.table.setItem(row, 1, QTableWidgetItem(status_text))
            self.table.setItem(row, 2, QTableWidgetItem(t["next_run"]))
            self.table.setItem(row, 3, QTableWidgetItem(t["last_run"]))
            self.table.setItem(row, 4, QTableWidgetItem(str(t["last_result"])))
    
            # 操作ボタン
            btn_widget = QWidget()
            btn_layout = QHBoxLayout(btn_widget)
            btn_layout.setContentsMargins(0,0,0,0)
    
            edit_btn = QPushButton("編集")
            edit_btn.clicked.connect(partial(self.edit_task, t["name"]))
            delete_btn = QPushButton("削除")
            delete_btn.clicked.connect(partial(self.delete_task, t["name"]))
            run_btn = QPushButton("今すぐ実行")
            run_btn.clicked.connect(partial(self.run_task, t["name"]))
    
            btn_layout.addWidget(edit_btn)
            btn_layout.addWidget(delete_btn)
            btn_layout.addWidget(run_btn)
            self.table.setCellWidget(row,5,btn_widget)

    def delete_task(self, task_name):
        reply = QMessageBox.question(
            self,
            "教官を削除",
            f"{task_name} を削除しますか？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
    
        def is_admin():
            try:
                return ctypes.windll.shell32.IsUserAnAdmin()
            except:
                return False
    
        # タスクの config パス取得
        deadline_id = task_name.replace(TASK_BASE_NAME + "_", "")
        config_path = get_task_config_path(deadline_id)
    
        # 管理者権限がない場合は昇格して再実行
        if not is_admin():
            script_path = os.path.abspath(__file__)
            args = f'"{script_path}" --delete "{config_path}"'
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, args, None, 1)
            return
    
        # 管理者権限ありなら直接削除
        try:
            import win32com.client
            service = win32com.client.Dispatch("Schedule.Service")
            service.Connect()
            root = service.GetFolder("\\")
            root.DeleteTask(task_name, 0)
            # 削除用の config ファイルも消す
            if os.path.exists(config_path):
                os.remove(config_path)
            QMessageBox.information(self, "削除完了", f"{task_name} を削除しました")
        except Exception as e:
            QMessageBox.warning(self, "削除失敗", f"削除に失敗しました:\n{e}")
    
        # 再描画
        self.load_tasks()


    def run_task(self, task_name):
        import win32com.client
        try:
            service = win32com.client.Dispatch("Schedule.Service")
            service.Connect()
            root = service.GetFolder("\\")
            task = root.GetTask(task_name)
            task.Run()
            QMessageBox.information(self, "実行完了", f"{task_name} を実行しました")
        except Exception as e:
            QMessageBox.warning(self, "実行失敗", f"実行に失敗しました:\n{e}")

    def edit_task(self, task_name):
        deadline_id = task_name.replace(TASK_BASE_NAME+"_","")
        config_path = get_task_config_path(deadline_id)
    
        if not os.path.exists(config_path):
            QMessageBox.warning(self,"エラー","設定ファイルが見つかりません")
            return
    
        with open(config_path,"r",encoding="utf-8") as f:
            cfg = json.load(f)
    
        dlg = TaskEditDialog(cfg,self)
    
        if dlg.exec() == QDialog.DialogCode.Accepted:
    
            # ★ ここで再読込するだけでOK
            with open(config_path,"r",encoding="utf-8") as f:
                updated_cfg = json.load(f)
    
            if hasattr(self,"main_app") and self.main_app:
                self.main_app.update_task(updated_cfg)
    
            QMessageBox.information(self,"保存完了",f"{updated_cfg.get('title','')} を更新しました")
            self.load_tasks()


class TaskEditDialog(QDialog):
    def __init__(self, cfg, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.setWindowTitle(f"タスク設定編集: {cfg.get('title','')}")
        self.resize(500, 500)
        layout = QVBoxLayout(self)

        # Excel
        layout.addWidget(QLabel("Excelファイル"))
        self.excel_input = QLineEdit(cfg.get("excel_path",""))
        browse_btn = QPushButton("参照")
        browse_btn.clicked.connect(self.browse_excel)
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_input)
        excel_layout.addWidget(browse_btn)
        layout.addLayout(excel_layout)

        # Webhook
        layout.addWidget(QLabel("Discord Webhook URL"))
        self.webhook_input = QLineEdit(cfg.get("webhook_url",""))
        layout.addWidget(self.webhook_input)

        # 日数・チェックボックス・時間など
        self.days_spin = QSpinBox()
        self.days_spin.setRange(0,60)
        self.days_spin.setValue(cfg.get("days_before_deadline",3))
        layout.addWidget(QLabel("締切何日前に通知"))
        layout.addWidget(self.days_spin)

        self.mention_checkbox = QCheckBox("メンションを有効（任意）")
        self.mention_checkbox.setChecked(cfg.get("mention_enabled",False))
        layout.addWidget(self.mention_checkbox)
        
        # 担当者（＋ーで追加）
        self.mention_box = QWidget()
        self.mention_layout = QVBoxLayout(self.mention_box)
        self.mention_layout.setContentsMargins(0,0,0,0)

        mentions = cfg.get("mentions", [])
        if mentions:
            for m in mentions:
                row = RowInput("担当名", "ユーザーID", self.mention_layout)
                row.short.setText(m.get("name",""))
                row.long.setText(m.get("id",""))
                self.mention_layout.addWidget(row)
        else:
            self.mention_layout.addWidget(RowInput("担当名", "ユーザーID", self.mention_layout, deletable=False))

        layout.addWidget(self.mention_box)

        for i in range(self.mention_layout.count()):
            w = self.mention_layout.itemAt(i).widget()
            if isinstance(w, RowInput):
                w.update_delete_state()
        
        layout.addWidget(QLabel("──────── 自動連絡 ────────"))
        self.auto_checkbox = QCheckBox("自動通知を有効（任意）")
        self.auto_checkbox.setChecked(cfg.get("auto_notify",False))
        layout.addWidget(self.auto_checkbox)

        self.time_edit = QTimeEdit(QTime.fromString(cfg.get("notify_time","09:00"),"HH:mm"))
        t_layout = QHBoxLayout()
        t_layout.addWidget(QLabel("通知時刻"))
        t_layout.addWidget(self.time_edit)
        layout.addLayout(t_layout)

        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1,30)
        self.interval_spin.setValue(cfg.get("notify_interval_days",1))
        i_layout = QHBoxLayout()
        i_layout.addWidget(QLabel("通知頻度（日）"))
        i_layout.addWidget(self.interval_spin)
        layout.addLayout(i_layout)

        save_btn = QPushButton("設定を保存")
        save_btn.clicked.connect(self.save)
        layout.addWidget(save_btn)

    def browse_excel(self):
        path,_ = QFileDialog.getOpenFileName(self,"Excelを選択","","Excel (*.xlsx *.xls)")
        if path:
            self.excel_input.setText(path)

    def save(self):
    
        # JSON 更新
        self.cfg["excel_path"] = self.excel_input.text()
        self.cfg["webhook_url"] = self.webhook_input.text()
        self.cfg["days_before_deadline"] = self.days_spin.value()
        self.cfg["mention_enabled"] = self.mention_checkbox.isChecked()
        self.cfg["auto_notify"] = self.auto_checkbox.isChecked()
        self.cfg["notify_time"] = self.time_edit.time().toString("HH:mm")
        self.cfg["notify_interval_days"] = self.interval_spin.value()
    
        # 担当者
        mentions = []
        for i in range(self.mention_layout.count()):
            w = self.mention_layout.itemAt(i).widget()
            if isinstance(w, RowInput):
                short, long = w.get()
                if short or long:
                    mentions.append({"name": short, "id": long})
        
        self.cfg["mentions"] = mentions
    
        # 保存
        config_path = get_config_path(self.cfg["deadline_id"])
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(self.cfg, f, ensure_ascii=False, indent=2)

        self.accept()

# ===============================
# 起動
# ===============================
if __name__ == "__main__":
    try:

        # ===== 通知モード =====
        if "--notify" in sys.argv:
            config_index = sys.argv.index("--notify") + 1
            config_path = sys.argv[config_index] if len(sys.argv) > config_index else None
            sys.exit(simekiri_notify.run_notify(config_path))

        # 管理者登録／削除モード
        if ADMIN_FLAG in sys.argv or '--delete' in sys.argv:
            config_path = None
            if '--delete' in sys.argv:
                idx = sys.argv.index('--delete') + 1
                if len(sys.argv) > idx:
                    config_path = sys.argv[idx]
            else:
                idx = sys.argv.index(ADMIN_FLAG) + 1
                if len(sys.argv) > idx:
                    config_path = sys.argv[idx]
        
            if not config_path or not os.path.exists(config_path):
                print("Config path missing")
                sys.exit(1)
        
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        
            task_name = get_task_name(cfg["deadline_id"])
        
            if '--delete' in sys.argv:
                import win32com.client
                service = win32com.client.Dispatch("Schedule.Service")
                service.Connect()
                root = service.GetFolder("\\")
                try:
                    root.DeleteTask(task_name, 0)
                    print(f"削除成功: {task_name}")
                except Exception as e:
                    print("削除失敗:", e)
                # configファイルも削除
                if os.path.exists(config_path):
                    os.remove(config_path)
                sys.exit(0)
        
            else:
                # 通常登録
                register_task_admin(cfg)
                sys.exit(0)


        # ===== 通常GUI起動 =====
        app = QApplication(sys.argv)
        win = NotifierApp()
        win.show()
        sys.exit(app.exec())

    except Exception:
        with open(ERROR_LOG, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
