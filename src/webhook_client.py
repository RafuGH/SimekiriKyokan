#webhook_client.py
#
# Webhook送信ヘルパー（Discord / Slack / Teams / Chatwork / Google Chat 自動判定）

import base64
from io import BytesIO

import requests


def detect_webhook_type(url: str) -> str:
    """URLからWebhookの種類を判定して返す。"""
    if not url:
        return "discord"
    if "hooks.slack.com" in url:
        return "slack"
    if "outlook.office.com" in url or "office365.com" in url or "webhook.office.com" in url:
        return "teams"
    if "chatwork.com" in url:
        return "chatwork"
    if "chat.googleapis.com" in url:
        return "googlechat"
    return "discord"


def build_mention(name: str, user_id: str, webhook_kind: str) -> str:
    """
    プラットフォーム別のメンション文字列を生成する。
    name    : 表示名（Teams / Chatwork で使用）
    user_id : ユーザーID / ユーザー名（Discord / Slack で使用）
    webhook_kind: detect_webhook_type() の戻り値

    各プラットフォームのメンション形式:
        Discord    : <@数字ID> または <@ユーザー名>
        Slack      : <@UXXXXXXXX>
        Teams      : <at>表示名</at>
        Chatwork   : [To:数字ID] 表示名
        Google Chat: @表示名のみ（メンション非対応）
    """
    if not user_id and not name:
        return ""
    if webhook_kind == "slack":
        return f"<@{user_id}>"
    elif webhook_kind == "teams":
        return f"<at>{name}</at>"
    elif webhook_kind == "chatwork":
        return f"[To:{user_id}] {name}"
    elif webhook_kind == "googlechat":
        return f"@{name}" if name else f"@{user_id}"
    else:
        # Discord: ユーザーIDが数字なら数字ID形式、そうでなければユーザー名形式
        return f"<@{user_id}>"


def send_webhook_text(url: str, content: str, embeds: list = None) -> requests.Response:
    """
    プラットフォームを自動判定してテキスト＋embed（任意）を送信する。
    対応: Discord / Slack / Teams / Chatwork / Google Chat
    """
    kind = detect_webhook_type(url)

    if kind == "slack":
        blocks = [{"type": "section", "text": {"type": "mrkdwn", "text": content}}]
        if embeds:
            for emb in embeds:
                title       = emb.get("title", "")
                desc        = emb.get("description", "")
                footer_text = emb.get("footer", {}).get("text", "")
                block_text  = f"*{title}*\n{desc}"
                if footer_text:
                    block_text += f"\n_{footer_text}_"
                blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": block_text}})
                blocks.append({"type": "divider"})
        return requests.post(url, json={"blocks": blocks})

    elif kind == "teams":
        body_items = [{"type": "TextBlock", "text": content, "wrap": True, "size": "Medium"}]
        if embeds:
            for emb in embeds:
                title       = emb.get("title", "")
                desc        = emb.get("description", "")
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

    elif kind == "chatwork":
        # Chatwork Incoming Webhook: body パラメータにテキストを送る
        # embed はテキストに展開して付加する
        lines = [content]
        if embeds:
            for emb in embeds:
                title       = emb.get("title", "")
                desc        = emb.get("description", "")
                footer_text = emb.get("footer", {}).get("text", "")
                if title:
                    lines.append(f"[title]{title}[/title]")
                if desc:
                    lines.append(desc)
                if footer_text:
                    lines.append(footer_text)
                lines.append("──────────")
        body_text = "\n".join(lines)
        return requests.post(url, data={"body": body_text})

    elif kind == "googlechat":
        # Google Chat Incoming Webhook: {"text": "..."} のみ対応
        # embed はテキストに展開して付加する
        lines = [content]
        if embeds:
            for emb in embeds:
                title       = emb.get("title", "")
                desc        = emb.get("description", "")
                footer_text = emb.get("footer", {}).get("text", "")
                if title:
                    lines.append(f"*{title}*")
                if desc:
                    lines.append(desc)
                if footer_text:
                    lines.append(f"_{footer_text}_")
                lines.append("──────────")
        return requests.post(url, json={"text": "\n".join(lines)})

    else:
        # Discord
        discord_payload = {"content": content}
        if embeds:
            discord_payload["embeds"] = embeds
        return requests.post(url, json=discord_payload)


def send_webhook_image(url: str, content: str, image_buffer: BytesIO) -> requests.Response:
    """
    プラットフォームを自動判定して画像を送信する。

    画像送信対応状況:
        Discord    : ✅ ファイル添付
        Slack      : ✗  Webhook では非対応（テキストのみ）
        Teams      : ✅ Base64 → Adaptive Card に埋め込み
        Chatwork   : ✗  Webhook では非対応（テキストのみ）
        Google Chat: ✅ Base64 → Card に埋め込み
    """
    kind = detect_webhook_type(url)

    if kind == "discord":
        return requests.post(
            url,
            files={"file": ("task.png", image_buffer, "image/png")},
            data={"content": content}
        )

    elif kind == "teams":
        # Teams: Adaptive Card に Base64 画像を埋め込む
        image_buffer.seek(0)
        image_b64 = base64.b64encode(image_buffer.read()).decode("utf-8")
        image_data_url = f"data:image/png;base64,{image_b64}"

        body_items = [{"type": "TextBlock", "text": content, "wrap": True}]
        body_items.append({
            "type": "Image",
            "url": image_data_url,
            "size": "Stretch"
        })

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

    elif kind == "googlechat":
        # Google Chat: Card に Base64 画像を埋め込む
        image_buffer.seek(0)
        image_b64 = base64.b64encode(image_buffer.read()).decode("utf-8")
        image_data_url = f"data:image/png;base64,{image_b64}"

        payload = {
            "cards": [{
                "header": {"title": "作業リスト"},
                "sections": [{
                    "widgets": [
                        {"textParagraph": {"text": content}},
                        {"image": {"imageUrl": image_data_url}}
                    ]
                }]
            }]
        }
        return requests.post(url, json=payload)

    else:
        # Slack, Chatwork など：テキストのみ送信
        return send_webhook_text(url, content)
