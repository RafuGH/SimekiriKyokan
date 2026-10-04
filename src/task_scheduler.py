#task_scheduler.py
#
# Windowsタスクスケジューラ連携・管理者権限まわりのヘルパー

import ctypes
import os
import sys
import uuid
import re
from datetime import datetime

APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
os.makedirs(APP_DIR, exist_ok=True)

TASK_BASE_NAME = "SimekiriKyokan"
ADMIN_FLAG = "--admin-register"
# 提出フォルダ監視用の定期タスク（締切通知タスクとは別。TASK_BASE_NAME で始めず一覧に混ざらないようにする）
WATCH_TASK_PREFIX = "SimekiriWatch_"
WATCH_INTERVAL = "PT10M"   # 10分ごと

ShellExecuteW = ctypes.windll.shell32.ShellExecuteW


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def get_config_path(deadline_id):
    return os.path.join(APP_DIR, f"{deadline_id}.json")


def get_task_name(deadline_id):
    return f"{TASK_BASE_NAME}_{deadline_id}"


def get_watch_task_name(deadline_id):
    return f"{WATCH_TASK_PREFIX}{deadline_id}"


def generate_deadline_id(category, end_date_str, title):
    safe_title = re.sub(r'[^a-zA-Z0-9ぁ-んァ-ン一-龯]', '', title)[:10]
    uid = uuid.uuid4().hex[:6]
    return f"{category}_{safe_title}_{uid}"


def check_schedule_dates(start, end):
    """
    自動通知の期間（datetime.date）を検証し、問題があればその文面を返す（問題なしは None）。
    終了日が過去だと、タスクスケジューラが EndBoundary エラーで登録を拒否する。
    """
    if end < datetime.now().date():
        return "終了日が過去の日付になっています。今日以降の日付を指定してください。"
    if end < start:
        return "終了日が開始日より前になっています。"
    return None


def relaunch_as_admin(config_path, extra_flag=ADMIN_FLAG):
    """管理者権限で自身を再起動し、タスク登録/削除を行わせる。"""
    if getattr(sys, 'frozen', False):
        ShellExecuteW(None, "runas", sys.executable, f'{extra_flag} "{config_path}"', None, 1)
    else:
        script_path = os.path.abspath(sys.argv[0])
        ShellExecuteW(None, "runas", sys.executable, f'"{script_path}" {extra_flag} "{config_path}"', None, 1)


def task_exists(deadline_id):
    try:
        import win32com.client
        service = win32com.client.Dispatch("Schedule.Service")
        service.Connect()
        service.GetFolder("\\").GetTask(get_task_name(deadline_id))
        return True
    except Exception:
        return False


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


def register_task_admin(config):
    import win32com.client
    task_name = get_task_name(config["deadline_id"])
    service = win32com.client.Dispatch("Schedule.Service")
    service.Connect()
    root = service.GetFolder("\\")
    try:
        root.DeleteTask(task_name, 0)
    except Exception:
        pass
    task_def = service.NewTask(0)
    h, m = map(int, config["notify_time"].split(":"))
    start_date = datetime.strptime(config["start_date"], "%Y-%m-%d")
    end_date   = datetime.strptime(config["end_date"], "%Y-%m-%d")
    start = start_date.replace(hour=h, minute=m, second=0)
    end   = end_date.replace(hour=23, minute=59, second=59)
    trigger = task_def.Triggers.Create(2)
    trigger.StartBoundary  = start.strftime("%Y-%m-%dT%H:%M:%S")
    trigger.EndBoundary    = end.strftime("%Y-%m-%dT%H:%M:%S")
    trigger.DaysInterval   = max(1, config["notify_interval_days"])
    trigger.Enabled        = True
    action = task_def.Actions.Create(0)
    if getattr(sys, 'frozen', False):
        action.Path = sys.executable
        config_path = get_config_path(config["deadline_id"])
        action.Arguments = f'--notify "{config_path}"'
        action.WorkingDirectory = os.path.dirname(sys.executable)
    else:
        action.Path = sys.executable
        config_path = get_config_path(config["deadline_id"])
        action.Arguments = f'"{os.path.abspath(sys.argv[0])}" --notify "{config_path}"'
        action.WorkingDirectory = os.path.dirname(os.path.abspath(sys.argv[0]))
    task_def.Principal.LogonType = 3
    task_def.Principal.RunLevel  = 0
    settings = task_def.Settings
    settings.Enabled            = True
    settings.StartWhenAvailable = True
    settings.ExecutionTimeLimit = "PT0S"
    root.RegisterTaskDefinition(task_name, task_def, 6, None, None, 3)
    config["task_registered"] = True
    _register_watch_task(service, root, config)


def delete_watch_task(deadline_id, root=None):
    """提出フォルダ監視タスクを削除する（無ければ何もしない）。"""
    try:
        if root is None:
            import win32com.client
            service = win32com.client.Dispatch("Schedule.Service")
            service.Connect()
            root = service.GetFolder("\\")
        root.DeleteTask(get_watch_task_name(deadline_id), 0)
    except Exception:
        pass


def _register_watch_task(service, root, config):
    """提出フォルダ監視が有効なら、一定間隔で --watch を実行するタスクを登録する。"""
    deadline_id = config["deadline_id"]
    delete_watch_task(deadline_id, root)
    if not config.get("submission_watch_enabled"):
        return

    import submission_watcher
    submission_watcher.reset_baseline(config)

    end_date = datetime.strptime(config["end_date"], "%Y-%m-%d").replace(hour=23, minute=59, second=59)
    start = max(datetime.now().replace(microsecond=0), datetime.strptime(config["start_date"], "%Y-%m-%d"))

    task_def = service.NewTask(0)
    trigger = task_def.Triggers.Create(1)   # 1 = 時刻トリガー
    trigger.StartBoundary = start.strftime("%Y-%m-%dT%H:%M:%S")
    trigger.EndBoundary = end_date.strftime("%Y-%m-%dT%H:%M:%S")
    trigger.Repetition.Interval = WATCH_INTERVAL
    trigger.Enabled = True

    action = task_def.Actions.Create(0)
    config_path = get_config_path(deadline_id)
    if getattr(sys, 'frozen', False):
        action.Path = sys.executable
        action.Arguments = f'--watch "{config_path}"'
        action.WorkingDirectory = os.path.dirname(sys.executable)
    else:
        action.Path = sys.executable
        action.Arguments = f'"{os.path.abspath(sys.argv[0])}" --watch "{config_path}"'
        action.WorkingDirectory = os.path.dirname(os.path.abspath(sys.argv[0]))
    task_def.Principal.LogonType = 3
    task_def.Principal.RunLevel = 0
    settings = task_def.Settings
    settings.Enabled = True
    settings.StartWhenAvailable = True
    settings.MultipleInstances = 2   # 前回の実行中なら新規起動しない
    settings.ExecutionTimeLimit = "PT10M"
    root.RegisterTaskDefinition(get_watch_task_name(deadline_id), task_def, 6, None, None, 3)
