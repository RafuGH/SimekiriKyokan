#task_image.py
#
# 締切通知に添付する作業リスト画像の生成

import os
from datetime import datetime
from io import BytesIO

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

STYLE_MAP = {
    "デザイナー": {"color": 0xFFD700},
    "プログラマー": {"color": 0x1E90FF},
    "サウンド": {"color": 0xFFA500},
    "未設定": {"color": 0x808080},
}

STATUS_COLOR_MAP = {
    "完了":     (180, 210, 255),
    "確認待ち": (180, 240, 200),
    "進行中":   (255, 245, 170),
    "未着手":   (245, 245, 245),
}

LEFT_ALIGN_COLUMNS = ["詳細", "備考"]


def pt_to_px(pt):
    return int(pt * 96 / 72)


def make_task_image(name, tasks, rate, column_order, col_width_map):
    """
    担当者ごとの作業一覧を表形式の画像(PNG)として生成する。
    column_order: Excel/Sheetsから読み込んだ列の並び順
    col_width_map: 列名 -> ピクセル幅 のマップ
    """
    DISPLAY_COLUMNS = [c for c in column_order if c in col_width_map]
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
    line_height = 20
    MAX_HEIGHT = 5000

    col_widths = [col_width_map[h] for h in headers]

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
