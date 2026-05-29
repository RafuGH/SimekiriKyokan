#simekiri_notify.py

import os
import sys
import json
import traceback
from datetime import datetime, timedelta

import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import textwrap

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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


# ===================================================
# Webhook送信ヘルパー（Discord / Slack / Teams 自動判定）
# ===================================================

def detect_webhook_type(url: str) -> str:
    """URLからWebhookの種類を判定して返す。"""
    if not url:
        return "discord"
    if "hooks.slack.com" in url:
        return "slack"
    if "outlook.office.com" in url or "office365.com" in url or "webhook.office.com" in url:
        return "teams"
    return "discord"


def send_webhook_text(url: str, content: str, embeds: list = None) -> requests.Response:
    """
    プラットフォームを自動判定してテキスト＋embed（任意）を送信する。
    Slack / Teams は embed を簡易テキストに変換して付加する。
    """
    kind = detect_webhook_type(url)

    if kind == "slack":
        # Slack: blocks API（シンプルにテキストのみ）
        blocks = [{"type": "section", "text": {"type": "mrkdwn", "text": content}}]
        if embeds:
            for emb in embeds:
                title = emb.get("title", "")
                desc  = emb.get("description", "")
                footer_text = emb.get("footer", {}).get("text", "")
                block_text = f"*{title}*\n{desc}"
                if footer_text:
                    block_text += f"\n_{footer_text}_"
                blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": block_text}})
                blocks.append({"type": "divider"})
        payload = {"blocks": blocks}
        return requests.post(url, json=payload)

    elif kind == "teams":
        # Teams: Adaptive Card (シンプルテキスト版)
        body_items = [{"type": "TextBlock", "text": content, "wrap": True, "size": "Medium"}]
        if embeds:
            for emb in embeds:
                title = emb.get("title", "")
                desc  = emb.get("description", "")
                footer_text = emb.get("footer", {}).get("text", "")
                if title:
                    body_items.append({"type": "TextBlock", "text": title, "weight": "Bolder", "wrap": True})
                if desc:
                    body_items.append({"type": "TextBlock", "text": desc, "wrap": True})
                if footer_text:
                    body_items.append({"type": "TextBlock", "text": footer_text, "isSubtle": True, "wrap": True})
                body_items.append({"type": "TextBlock", "text": "──────────", "isSubtle": True})

        payload = {
            "type": "message",
            "attachments": [{
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.2",
                    "body": body_items
                }
            }]
        }
        return requests.post(url, json=payload)

    else:
        # Discord（従来通り）
        discord_payload = {"content": content}
        if embeds:
            discord_payload["embeds"] = embeds
        return requests.post(url, json=discord_payload)


def send_webhook_image(url: str, content: str, image_buffer: BytesIO) -> requests.Response:
    """
    プラットフォームを自動判定して画像を送信する。
    Slack / Teams は画像添付非対応のため、テキストのみ送信する。
    """
    kind = detect_webhook_type(url)

    if kind == "discord":
        return requests.post(
            url,
            files={"file": ("task.png", image_buffer, "image/png")},
            data={"content": content}
        )
    else:
        # Slack / Teams: 画像送信は Webhook 非対応のためテキストのみ
        return send_webhook_text(url, content)


# ===================================================
# Google Sheets 読み込みヘルパー
# ===================================================

def load_dataframe_from_sheets(spreadsheet_url: str, credentials_path: str, sheet_name: str = "作業リスト"):
    """
    Google Sheets APIでスプレッドシートを読み込み DataFrameを返す。
    戻り値: (df, col_width_map, row_height_base)
    ※列幅・行高さはデフォルト値を使用する（Sheets APIでは取得不可）。
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as e:
        raise ImportError(
            "gspread または google-auth がインストールされていません。\n"
            "pip install gspread google-auth を実行してください。\n" + str(e)
        )

    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    gc = gspread.authorize(creds)

    # URLからスプレッドシートを開く
    sh = gc.open_by_url(spreadsheet_url)
    ws = sh.worksheet(sheet_name)

    all_values = ws.get_all_values()

    if not all_values:
        raise ValueError("スプレッドシートにデータがありません")

    # C列（index=2）から K列（index=10）に相当する列を取得
    # ヘッダー行を特定（空でない最初の行）
    header_row_idx = 0
    for i, row in enumerate(all_values):
        if any(cell.strip() for cell in row[2:11]):
            header_row_idx = i
            break

    headers = all_values[header_row_idx][2:11]
    data_rows = [r[2:11] for r in all_values[header_row_idx + 1:] if any(c.strip() for c in r[2:11])]

    df = pd.DataFrame(data_rows, columns=headers)

    # デフォルト列幅マップ（Sheets APIでは列幅取得不可のためデフォルト値）
    DEFAULT_COL_WIDTH = 120
    col_width_map = {h: DEFAULT_COL_WIDTH for h in headers}
    # よく使う列は少し広めに
    for col, w in {"内容": 160, "詳細": 200, "備考": 140, "担当": 80, "締切": 70}.items():
        if col in col_width_map:
            col_width_map[col] = w

    row_height_base = 24  # デフォルト行高さ（px）

    return df, col_width_map, row_height_base


# ===================================================
# メイン処理
# ===================================================

def run_notify(config_path=None, test_mode=False):
    def pt_to_px(pt):
        return int(pt * 96 / 72)

    if getattr(sys, 'frozen', False):
        BASE_DIR = os.path.dirname(sys.executable)
    else:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
    os.makedirs(APP_DIR, exist_ok=True)

    global LOG_FILE
    LOG_FILE = os.path.join(APP_DIR, "simekiri_run_log.txt")

    print("ARGV:", sys.argv)
    print("CONFIG_PATH:", config_path)

    if not config_path:
        print("ERROR: config_path is None")
        return 1

    CONFIG_FILE = config_path
    write_log(f"Using config file: {CONFIG_FILE}")

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)

    WEBHOOK_URL = config.get("webhook_url")
    if not WEBHOOK_URL:
        write_log("Webhook URL is missing")
        return 1

    DAYS_BEFORE = config.get("days_before_deadline", 3)
    MENTION_ENABLED = config.get("mention_enabled", False)
    MENTION_MAP = config.get("mentions", {})

    # ---- 確認待ち通知先（レビュアー）設定 ----
    REVIEWER_ENABLED = config.get("reviewer_enabled", False)
    REVIEWER_WEBHOOK = config.get("reviewer_webhook_url", "") or WEBHOOK_URL
    REVIEWER_MENTION_MAP = config.get("reviewer_mentions", {})
    # list → dict 変換（レビュアー）
    if isinstance(REVIEWER_MENTION_MAP, list):
        fixed = {}
        for item in REVIEWER_MENTION_MAP:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).strip()
            user_id = str(item.get("id", "")).strip()
            if name and user_id:
                fixed[name] = f"<@{user_id}>" if detect_webhook_type(REVIEWER_WEBHOOK) == "discord" else f"<@{user_id}>"
        REVIEWER_MENTION_MAP = fixed

    # list → dict 変換（通常メンション）
    if isinstance(MENTION_MAP, list):
        write_log("mentions is list → converting to dict")
        fixed = {}
        for item in MENTION_MAP:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).strip()
            user_id = str(item.get("id", "")).strip()
            if name and user_id:
                fixed[name] = f"<@{user_id}>"
        MENTION_MAP = fixed

    write_log(f"MENTION_ENABLED={MENTION_ENABLED}")
    write_log(f"MENTION_MAP={MENTION_MAP}")
    write_log(f"REVIEWER_ENABLED={REVIEWER_ENABLED}")

    IS_TEST = test_mode

    if IS_TEST:
        write_log("=== TEST MODE ===")
        try:
            r = send_webhook_text(WEBHOOK_URL, "🧪 **締切教官 通知テスト**\nこのメッセージが見えていれば正常です。")
            status = getattr(r, "status_code", None)
            write_log(f"Test notify sent. status={status}")
            if status in (200, 204):
                return 0
            else:
                return 1
        except Exception as e:
            write_log("Test notify failed: " + repr(e))
            return 1

    def _choose_log_dir():
        try:
            os.makedirs(APP_DIR, exist_ok=True)
            return APP_DIR
        except Exception:
            return os.path.expanduser("~")

    LOG_DIR = _choose_log_dir()
    LOG_FILE = os.path.join(LOG_DIR, "simekiri_run_log.txt")
    ERR_FILE = os.path.join(LOG_DIR, "simekiri_error_log.txt")

    write_log("=== start run pid=" + str(os.getpid()) + " cwd=" + os.getcwd() + " ===")

    try:
        APP_DIR = os.path.join(os.environ["LOCALAPPDATA"], "SimekiriKyokan")
        os.makedirs(APP_DIR, exist_ok=True)

        EXCEL_FILE    = config.get("excel_path", "")
        SHEETS_URL    = config.get("sheets_url", "")
        SHEETS_CREDS  = config.get("sheets_credentials_path", "")
        DATA_SOURCE   = config.get("data_source", "excel")  # "excel" or "sheets"

        WEBHOOK_URL           = config.get("webhook_url", "")
        DAYS_BEFORE_DEADLINE  = config.get("days_before_deadline", 3)
        MENTION_ENABLED       = config.get("mention_enabled", False)

        # -------------------------
        # Helpers: date conversions
        # -------------------------
        def convert_deadline_value(x):
            if pd.isna(x):
                return pd.NaT
            if isinstance(x, (int, float)):
                try:
                    return (pd.to_datetime("1899-12-30") + pd.to_timedelta(x, unit="D"))
                except Exception:
                    return pd.to_datetime(x, errors="coerce")
            try:
                return pd.to_datetime(x, errors="coerce")
            except Exception:
                return pd.NaT

        # -------------------------
        # データ読み込み（Excel or Google Sheets）
        # -------------------------
        COL_WIDTH_MAP = {}
        row_height_base = pt_to_px(15)

        if DATA_SOURCE == "sheets":
            # ---------- Google Sheets ----------
            if not SHEETS_URL:
                write_log("sheets_url is missing")
                return 1
            if not SHEETS_CREDS or not os.path.exists(SHEETS_CREDS):
                write_log(f"sheets_credentials_path が見つかりません: {SHEETS_CREDS}")
                return 1

            write_log(f"Google Sheets から読み込み中: {SHEETS_URL}")
            try:
                df, COL_WIDTH_MAP, row_height_base = load_dataframe_from_sheets(
                    SHEETS_URL, SHEETS_CREDS
                )
            except Exception as e:
                write_log("Failed to read Google Sheets: " + repr(e))
                raise

        else:
            # ---------- Excel ----------
            if not EXCEL_FILE or not os.path.exists(EXCEL_FILE):
                write_log(f"Excel NOT FOUND: {EXCEL_FILE}")
                return 1

            try:
                df = pd.read_excel(
                    EXCEL_FILE,
                    sheet_name="作業リスト",
                    usecols="C:K"
                )

                wb = load_workbook(EXCEL_FILE)
                ws = wb["作業リスト"]

                excel_width_map = {}
                start_col_index = 3

                for i, col_name in enumerate(df.columns):
                    excel_col_index = start_col_index + i
                    letter = get_column_letter(excel_col_index)
                    dim = ws.column_dimensions.get(letter)
                    if dim and dim.width:
                        pixel_width = int(dim.width * 8.2 + 12)
                        excel_width_map[col_name] = pixel_width
                    else:
                        excel_width_map[col_name] = 120

                COL_WIDTH_MAP = excel_width_map

                data_row_index = 8
                excel_row_height = ws.row_dimensions[data_row_index].height
                if excel_row_height:
                    row_height_base = int(excel_row_height * 96 / 72) + 3
                else:
                    row_height_base = pt_to_px(15)

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
        today_str = today.strftime("%Y%m%d")

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

        STYLE_MAP = {
            "デザイナー": {"color": 0xFFD700},
            "プログラマー": {"color": 0x1E90FF},
            "サウンド": {"color": 0xFFA500},
            "未設定": {"color": 0x808080},
        }

        WRAP_RULES = {
            "職種": 8, "分類": 10, "内容": 14, "詳細": 34,
            "担当": 4, "進捗": 4, "優先度": 4, "備考": 10, "締切": 6,
        }

        def make_task_image(name, tasks, rate):
            DISPLAY_COLUMNS = [c for c in COLUMN_ORDER if c in COL_WIDTH_MAP]
            headers = DISPLAY_COLUMNS

            font_path = os.path.join(os.environ["WINDIR"], "Fonts", "meiryo.ttc")

            try:
                title_font  = ImageFont.truetype(font_path, pt_to_px(16))
                header_font = ImageFont.truetype(font_path, pt_to_px(11))
                text_font   = ImageFont.truetype(font_path, pt_to_px(11))
            except Exception:
                title_font  = ImageFont.load_default()
                header_font = ImageFont.load_default()
                text_font   = ImageFont.load_default()

            TOP_PADDING = 6
            LEFT_PADDING = 10
            LEFT_ALIGN_COLUMNS = ["詳細", "備考"]
            line_height = 20
            MAX_HEIGHT = 5000

            STATUS_COLOR_MAP = {
                "完了":     (180, 210, 255),
                "確認待ち": (180, 240, 200),
                "進行中":   (255, 245, 170),
                "未着手":   (245, 245, 245),
            }

            col_widths = [COL_WIDTH_MAP[h] for h in headers]

            def wrap_text_pixel(text, max_width):
                if not text:
                    return [""]
                dummy_img = Image.new("RGB", (1, 1))
                draw_dummy = ImageDraw.Draw(dummy_img)
                lines = []
                for raw_line in str(text).splitlines():
                    current = ""
                    for char in raw_line:
                        if draw_dummy.textlength(current + char, font=text_font) <= max_width - 12:
                            current += char
                        else:
                            lines.append(current)
                            current = char
                    lines.append(current)
                return lines

            wrapped_rows = []
            for _, row in tasks.iterrows():
                dl = row["締切"]
                try:
                    deadline_date = dl.date() if hasattr(dl, "date") else pd.to_datetime(dl).date()
                except Exception:
                    deadline_date = datetime.now().date()
                deadline_text = deadline_date.strftime("%m/%d")

                values = []
                for h in headers:
                    if h == "締切":
                        values.append(deadline_text)
                    elif h == "進捗":
                        status = str(row.get("進捗_raw", "")).strip()
                        status_icon_map = {"完了": "完了", "確認待ち": "確認待ち", "進行中": "進行中", "未着手": "未着手"}
                        values.append(status_icon_map.get(status, status))
                    else:
                        values.append(row.get(h, ""))

                wrapped = [wrap_text_pixel(val, col_widths[i]) for i, val in enumerate(values)]
                max_lines = max(len(cell) for cell in wrapped)
                status_raw = str(row.get("進捗_raw", "")).strip()
                wrapped_rows.append((wrapped, max_lines, status_raw))

            header_height = 140
            total_height = header_height + sum((max_lines * line_height + TOP_PADDING*2) for _, max_lines, _ in wrapped_rows) + 40 + 80
            total_height = min(total_height, MAX_HEIGHT)
            total_width  = sum(col_widths) + 40

            img  = Image.new("RGB", (total_width, total_height), "white")
            draw = ImageDraw.Draw(img)

            title = f"{name} の作業（完了率 {rate}%）"
            try:
                title_w = draw.textbbox((0,0), title, font=title_font)[2]
            except Exception:
                title_w = draw.textlength(title, font=title_font)
            draw.text(((total_width - title_w)/2, 20), title, fill="black", font=title_font)

            y = 90
            x_start = 20
            x = x_start
            for i, header in enumerate(headers):
                draw.rectangle([x, y, x + col_widths[i], y + 45], fill=(230,230,230), outline="black", width=1)
                text_w = draw.textlength(header, font=header_font)
                draw.text((x + (col_widths[i]-text_w)/2, y+10), header, fill="black", font=header_font)
                x += col_widths[i]
            y += 45

            for wrapped, max_lines, status_raw in wrapped_rows:
                row_height = max_lines * line_height + TOP_PADDING*2
                x = x_start
                bg_color = STATUS_COLOR_MAP.get(status_raw, (255,255,255))
                for col_index, (col_name, cell_lines) in enumerate(zip(headers, wrapped)):
                    w = col_widths[col_index]
                    draw.rectangle([x, y, x + w, y + row_height], fill=bg_color, outline="black", width=1)
                    total_cell_height = len(cell_lines)*line_height
                    if col_name in LEFT_ALIGN_COLUMNS:
                        start_y = y + TOP_PADDING
                    else:
                        start_y = y + (row_height - total_cell_height)/2
                    for i, line in enumerate(cell_lines):
                        line_y = start_y + i*line_height
                        if col_name in LEFT_ALIGN_COLUMNS:
                            draw.text((x + LEFT_PADDING, line_y), line, font=text_font, fill="black")
                        else:
                            line_w = draw.textlength(line, font=text_font)
                            draw.text((x + (w - line_w)/2, line_y), line, font=text_font, fill="black")
                    x += w
                y += row_height

            buffer = BytesIO()
            img.save(buffer, format="PNG")
            buffer.seek(0)
            return buffer

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
                image_buffer = make_task_image(namekey, group, rate)
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

            embed = {
                "title": f"📋 {namekey}の作業一覧（完了率 {rate}%）",
                "description": description_text,
                "color": embed_color,
                "footer": {"text": f"更新日: {today.strftime('%Y/%m/%d')}"}
            }
            all_embeds.append(embed)

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

            reviewer_embeds = []

            for 担当表示, group in review_pending.groupby("担当"):
                namekey = str(担当表示).strip() if 担当表示 is not None else "未設定"

                # レビュアーメンション取得
                reviewer_mention = ""
                if isinstance(REVIEWER_MENTION_MAP, dict):
                    reviewer_mention = REVIEWER_MENTION_MAP.get(namekey, "")

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

                reviewer_embed = {
                    "title": f"🔍 確認待ちタスク（担当: {namekey}）",
                    "description": description_text,
                    "color": 0x2ECC71,
                    "footer": {"text": f"更新日: {today.strftime('%Y/%m/%d')}"}
                }
                reviewer_embeds.append(reviewer_embed)

                # 画像送信（担当者ごと）
                try:
                    img_buf = make_task_image(namekey, group, person_rates.get(namekey, 0))
                    r_img = send_webhook_image(
                        REVIEWER_WEBHOOK,
                        f"🔍 {reviewer_mention} {namekey} の確認待ち作業リスト",
                        img_buf
                    )
                    write_log(f"Posted reviewer image for {namekey}, status={getattr(r_img,'status_code','N/A')}")
                except Exception as e:
                    write_log(f"Failed to post reviewer image for {namekey}: {repr(e)}")

            reviewer_embeds = reviewer_embeds[:10]

            try:
                r_rev = send_webhook_text(
                    REVIEWER_WEBHOOK,
                    f"🔍 **確認待ちタスクのお知らせ**\n以下のタスクが確認待ち状態です。ご確認をお願いします。",
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