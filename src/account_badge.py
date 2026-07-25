#account_badge.py
#
# タイトルバーに表示するGoogleアカウント状態バッジ

import base64
import json as _json

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont, QPainter, QBrush, QPen, QPixmap
from PyQt6.QtWidgets import QWidget, QHBoxLayout, QLabel

import google_auth_helper


class AccountBadge(QWidget):
    """タイトルバー右端に表示するアカウントバッジ"""
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 4, 0)
        layout.setSpacing(6)

        self.icon_lbl = QLabel()
        self.icon_lbl.setFixedSize(22, 22)
        self.name_lbl = QLabel()
        self.name_lbl.setObjectName("desc_lbl")

        layout.addWidget(self.icon_lbl)
        layout.addWidget(self.name_lbl)
        self.refresh()

    def refresh(self):
        if google_auth_helper.has_token():
            display = "認証済み"
            try:
                with open(google_auth_helper.TOKEN_PATH, "r") as f:
                    data = _json.load(f)
                id_tok = data.get("id_token", "")
                if id_tok:
                    payload = id_tok.split(".")[1]
                    payload += "=" * (-len(payload) % 4)
                    info = _json.loads(base64.b64decode(payload))
                    display = info.get("email", display)
            except Exception:
                pass

            self._draw_icon(True)
            self.name_lbl.setText(display)
            self.setToolTip(f"Google アカウント: {display}")
        else:
            self._draw_icon(False)
            self.name_lbl.setText("未ログイン")
            self.setToolTip("Google Sheets を使う場合は認可が必要です")

    def _draw_icon(self, logged_in: bool):
        pix = QPixmap(22, 22)
        pix.fill(Qt.GlobalColor.transparent)
        p = QPainter(pix)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        color = QColor("#107c10") if logged_in else QColor("#797979")
        p.setBrush(QBrush(color))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(0, 0, 22, 22)
        p.setPen(QPen(QColor("white")))
        p.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        p.drawText(pix.rect(), Qt.AlignmentFlag.AlignCenter, "G" if logged_in else "—")
        p.end()
        self.icon_lbl.setPixmap(pix)
