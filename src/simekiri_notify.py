#simekiri_notify.py
#
# 締切通知のオーケストレーション（データ読み込み → 集計 → Webhook送信）

import os
import sys
import json
import traceback
from datetime import datetime

import pandas as pd

from webhook_client import detect_webhook_type, build_mention, send_webhook_text, send_webhook_image
from task_image import make_task_image, STYLE_MAP
from data_loader import convert_deadline_value, load_dataframe_from_excel, load_dataframe_from_sheets

LOG_FILE = None


def write_log(message):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if LOG_FILE:
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"[{now}] {message}\n")
        except Exception:
            pass

    print(f"[{now}] {message}")


def _mentions_list_to_dict(mention_list, webhook_kind):
    """
    [{"name": "...", "id": "..."}] 形式のメンション設定を
    {"担当名": "<@ID>"} 形式（プラットフォーム別文字列）に変換する。
    """
    fixed = {}
    for item in mention_list:
        if not isinstance(item, dict):
            continue
        name    = str(item.get("name", "")).strip()
        user_id = str(item.get("id", "")).strip()
        if not (name or user_id):
            continue
        key = name if name else user_id
        fixed[key] = build_mention(name, user_id, webhook_kind)
    return fixed


def _choose_log_dir(app_dir):
    try:
        os.makedirs(app_dir, exist_ok=True)
        return app_dir
    except Exception:
        return os.path.expanduser("~")


def run_notify(config_path=None, test_mode=False):
    global LOG_FILE

    APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
    LOG_DIR = _choose_log_dir(APP_DIR)
    LOG_FILE = os.path.join(LOG_DIR, "simekiri_run_log.txt")
    ERR_FILE = os.path.join(LOG_DIR, "simekiri_error_log.txt")

    print("ARGV:", sys.argv)
    print("CONFIG_PATH:", config_path)

    if not config_path:
        print("ERROR: config_path is None")
        return 1

    write_log(f"Using config file: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    WEBHOOK_URL = config.get("webhook_url", "")
    if not WEBHOOK_URL:
        write_log("Webhook URL is missing")
        return 1

    DAYS_BEFORE_DEADLINE = config.get("days_before_deadline", 3)
    MENTION_ENABLED      = config.get("mention_enabled", False)
    MENTION_MAP          = config.get("mentions", {})

    # ---- 確認待ち通知先（レビュアー）設定 ----
    REVIEWER_ENABLED     = config.get("reviewer_enabled", False)
    REVIEWER_WEBHOOK     = config.get("reviewer_webhook_url", "") or WEBHOOK_URL
    REVIEWER_MENTION_MAP = config.get("reviewer_mentions", {})

    _reviewer_webhook_kind = detect_webhook_type(REVIEWER_WEBHOOK)
    if isinstance(REVIEWER_MENTION_MAP, list):
        REVIEWER_MENTION_MAP = _mentions_list_to_dict(REVIEWER_MENTION_MAP, _reviewer_webhook_kind)
    write_log(f"REVIEWER_MENTION_MAP={REVIEWER_MENTION_MAP}")

    _main_webhook_kind = detect_webhook_type(WEBHOOK_URL)
    if isinstance(MENTION_MAP, list):
        write_log("mentions is list → converting to dict")
        MENTION_MAP = _mentions_list_to_dict(MENTION_MAP, _main_webhook_kind)

    write_log(f"MENTION_ENABLED={MENTION_ENABLED}")
    write_log(f"MENTION_MAP={MENTION_MAP}")
    write_log(f"REVIEWER_ENABLED={REVIEWER_ENABLED}")

    if test_mode:
        write_log("=== TEST MODE ===")
        try:
            r = send_webhook_text(WEBHOOK_URL, "🧪 **締切教官 通知テスト**\nこのメッセージが見えていれば正常です。")
            status = getattr(r, "status_code", None)
            write_log(f"Test notify sent. status={status}")
            return 0 if status in (200, 204) else 1
        except Exception as e:
            write_log("Test notify failed: " + repr(e))
            return 1

    write_log("=== start run pid=" + str(os.getpid()) + " cwd=" + os.getcwd() + " ===")

    try:
        EXCEL_FILE  = config.get("excel_path", "")
        SHEETS_URL  = config.get("sheets_url", "")
        DATA_SOURCE = config.get("data_source", "excel")  # "excel" or "sheets"

        # -------------------------
        # データ読み込み（Excel or Google Sheets）
        # -------------------------
        if DATA_SOURCE == "sheets":
            if not SHEETS_URL:
                write_log("sheets_url is missing")
                return 1

            write_log(f"Google Sheets から読み込み中: {SHEETS_URL}")
            try:
                df, COL_WIDTH_MAP, row_height_base = load_dataframe_from_sheets(SHEETS_URL)
            except Exception as e:
                write_log("Failed to read Google Sheets: " + repr(e))
                raise

        else:
            if not EXCEL_FILE or not os.path.exists(EXCEL_FILE):
                write_log(f"Excel NOT FOUND: {EXCEL_FILE}")
                return 1

            try:
                df, COL_WIDTH_MAP, row_height_base = load_dataframe_from_excel(EXCEL_FILE)
            except Exception as e:
                write_log("Failed to read excel: " + repr(e))
                raise

        write_log("COLUMNS: " + str(list(df.columns)))

        REQUIRED_COLUMNS = ["内容", "締切", "担当", "進捗"]
        COLUMN_ORDER = list(df.columns)

        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            write_log(f"Missing columns: {missing}")
            return 1

        df = df[df["担当"].notna()]
        df = df[df["担当"].astype(str).str.strip() != ""]

        # いらない列を除外
        df.columns = df.columns.str.strip()
        df = df.drop(columns=[c for c in df.columns if "Unnamed" in c or c == "目次"], errors="ignore")

        def convert_done(v):
            return str(v).strip() == "完了"

        df["進捗_raw"] = df["進捗"]
        df["進捗"] = df["進捗"].apply(convert_done)

        df["締切"] = df["締切"].apply(convert_deadline_value)
        df = df.dropna(subset=["内容", "締切"])

        if not pd.api.types.is_datetime64_any_dtype(df["締切"]):
            df["締切"] = pd.to_datetime(df["締切"], errors="coerce")

        today = datetime.now().date()

        total_tasks = len(df)
        completed_tasks = int(df["進捗"].sum()) if "進捗" in df else 0
        overall_rate = round((completed_tasks / total_tasks) * 100) if total_tasks > 0 else 0

        person_rates = {}
        if "担当" in df.columns:
            for person, g in df[df["担当"].notna()].groupby("担当"):
                total = len(g)
                done = int(g["進捗"].sum())
                person_rates[str(person).strip()] = round((done / total) * 100) if total > 0 else 0

        df["days_left"] = (df["締切"].dt.date - today)
        df["days_left"] = df["days_left"].apply(lambda x: x.days if pd.notna(x) else 9999)

        if "優先度" in df.columns:
            not_unnecessary = ~df["優先度"].astype(str).str.strip().isin(["不要"])
        else:
            not_unnecessary = True

        # 確認待ちは担当者通知から除外し、レビュアーにのみ通知する
        is_review_waiting = df["進捗_raw"].astype(str).str.strip() == "確認待ち" \
            if "進捗_raw" in df.columns else pd.Series(False, index=df.index)

        pending = df[
            (df["進捗"] == False) &
            (~is_review_waiting) &          # ← 確認待ちを除外
            not_unnecessary &
            (df["days_left"] <= DAYS_BEFORE_DEADLINE)
        ]

        # ---- 確認待ちタスク（締切に関わらず全件・レビュアー専用） ----
        review_pending = df[is_review_waiting] if "進捗_raw" in df.columns else pd.DataFrame()

        if pending.empty and review_pending.empty:
            try:
                r = send_webhook_text(WEBHOOK_URL, "🎉 締切の近い、または過ぎた作業は現在ありません。")
                write_log(f"No pending tasks. webhook status: {getattr(r,'status_code', 'N/A')}")
            except Exception as e:
                write_log("Webhook send failed (no pending): " + repr(e))
            return 0

        # -------------------------
        # 通常の締切通知（担当者ごと）
        # -------------------------
        all_embeds = []
        errors = []

        for 担当表示, group in pending.groupby("担当"):
            namekey = str(担当表示).strip() if 担当表示 is not None else "未設定"
            rate = person_rates.get(namekey, 0)

            mention = ""
            if MENTION_ENABLED and isinstance(MENTION_MAP, dict):
                mention = MENTION_MAP.get(namekey, "")

            try:
                image_buffer = make_task_image(namekey, group, rate, COLUMN_ORDER, COL_WIDTH_MAP)
                r = send_webhook_image(
                    WEBHOOK_URL,
                    f"📗 {mention} {namekey} の作業リスト",
                    image_buffer
                )
                write_log(f"Posted image for {namekey}, status={getattr(r,'status_code','N/A')}")
            except Exception as e:
                write_log(f"Failed to post image for {namekey}: {repr(e)}")
                errors.append((namekey, repr(e)))

            representative_job = str(group.iloc[0].get("職種", "未設定")).strip()
            embed_color = STYLE_MAP.get(representative_job, STYLE_MAP["未設定"])["color"]

            lines = []
            for _, row in group.iterrows():
                try:
                    days_left = (row["締切"].date() - today).days
                except Exception:
                    days_left = 0

                if days_left < 0:
                    days_text = f"🔴 締切が {abs(days_left)} 日過ぎてる！"
                elif days_left == 0:
                    days_text = "🟠 今日が締切！"
                elif days_left == 1:
                    days_text = "🟡 明日が締切！"
                elif days_left <= 3:
                    days_text = f"🟡 締切まであと {days_left} 日！"
                else:
                    days_text = f"⚪ 締切まであと {days_left} 日"

                lines.append(f"・{row['内容']}（{days_text}）")

            description_text = "\n".join(lines)
            if len(description_text) > 4000:
                description_text = description_text[:3900] + "\n…（以下省略）"

            all_embeds.append({
                "title": f"📋 {namekey}の作業一覧（完了率 {rate}%）",
                "description": description_text,
                "color": embed_color,
                "footer": {"text": f"更新日: {today.strftime('%Y/%m/%d')}"}
            })

        all_embeds = all_embeds[:10]

        if DAYS_BEFORE_DEADLINE == 0:
            deadline_text = "今日が締切の作業、または締切を過ぎた作業があるぞ！"
        elif DAYS_BEFORE_DEADLINE == 1:
            deadline_text = "明日が締切の作業、または締切を過ぎた作業があるぞ！"
        else:
            deadline_text = f"{DAYS_BEFORE_DEADLINE}日以内に締切の作業、または締切を過ぎた作業があるぞ！"

        if not pending.empty:
            summary_content = (
                f"⚠️ **本日の締切連絡** ⚠️\n"
                f"✅ 全体完了率：{overall_rate}%\n"
                f"{deadline_text}"
            )
            try:
                r2 = send_webhook_text(WEBHOOK_URL, summary_content, all_embeds)
                write_log(f"Posted summary payload status={getattr(r2,'status_code','N/A')}")
                if getattr(r2, "status_code", 0) >= 400:
                    write_log("Summary post failed: " + (r2.text if hasattr(r2, "text") else ""))
            except Exception as e:
                write_log("Failed to post summary payload: " + repr(e))

        # -------------------------
        # 確認待ちタスクをレビュアーに通知
        # -------------------------
        if REVIEWER_ENABLED and not review_pending.empty:
            write_log(f"確認待ちタスク {len(review_pending)} 件をレビュアーに通知します")

            # レビュアーメンションを全員分まとめる（確認待ち通知は「レビュアー全員」に飛ばす）
            all_reviewer_mentions = ""
            if isinstance(REVIEWER_MENTION_MAP, dict) and REVIEWER_MENTION_MAP:
                all_reviewer_mentions = " ".join(REVIEWER_MENTION_MAP.values())
            write_log(f"reviewer mentions: {all_reviewer_mentions!r}")

            reviewer_embeds = []

            for 担当表示, group in review_pending.groupby("担当"):
                namekey = str(担当表示).strip() if 担当表示 is not None else "未設定"

                lines = []
                for _, row in group.iterrows():
                    try:
                        dl_date = row["締切"].date() if hasattr(row["締切"], "date") else pd.to_datetime(row["締切"]).date()
                        deadline_str = dl_date.strftime("%m/%d")
                    except Exception:
                        deadline_str = "不明"
                    lines.append(f"・{row['内容']}（締切: {deadline_str}｜担当: {namekey}）")

                description_text = "\n".join(lines)
                if len(description_text) > 4000:
                    description_text = description_text[:3900] + "\n…（以下省略）"

                reviewer_embeds.append({
                    "title": f"🔍 確認待ちタスク（担当: {namekey}）",
                    "description": description_text,
                    "color": 0x2ECC71,
                    "footer": {"text": f"更新日: {today.strftime('%Y/%m/%d')}"}
                })

                # 画像送信（担当者ごと）
                try:
                    img_buf = make_task_image(namekey, group, person_rates.get(namekey, 0), COLUMN_ORDER, COL_WIDTH_MAP)
                    r_img = send_webhook_image(
                        REVIEWER_WEBHOOK,
                        f"🔍 {namekey} の確認待ち作業リスト",
                        img_buf
                    )
                    write_log(f"Posted reviewer image for {namekey}, status={getattr(r_img,'status_code','N/A')}")
                except Exception as e:
                    write_log(f"Failed to post reviewer image for {namekey}: {repr(e)}")

            reviewer_embeds = reviewer_embeds[:10]

            mention_prefix = f"{all_reviewer_mentions}\n" if all_reviewer_mentions else ""
            try:
                r_rev = send_webhook_text(
                    REVIEWER_WEBHOOK,
                    f"{mention_prefix}🔍 **確認待ちタスクのお知らせ**\n以下のタスクが確認待ち状態です。ご確認をお願いします。",
                    reviewer_embeds
                )
                write_log(f"Posted reviewer summary status={getattr(r_rev,'status_code','N/A')}")
            except Exception as e:
                write_log("Failed to post reviewer summary: " + repr(e))

        if errors:
            write_log("Some posting errors: " + repr(errors))

        write_log("正常終了")
        return 0

    except Exception:
        try:
            with open(ERR_FILE, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except Exception:
            pass
        write_log("EXCEPTION: see " + ERR_FILE)
        return 1
