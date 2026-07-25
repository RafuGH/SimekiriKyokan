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


def generate_deadline_id(category, end_date_str, title):
    safe_title = re.sub(r'[^a-zA-Z0-9ぁ-んァ-ン一-龯]', '', title)[:10]
    uid = uuid.uuid4().hex[:6]
    return f"{category}_{safe_title}_{uid}"


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
