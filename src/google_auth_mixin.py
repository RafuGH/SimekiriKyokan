#google_auth_mixin.py
#
# NotifierApp / TaskEditDialog に「Googleで認可する」ボタンの挙動を提供するミックスイン。
# 要件: ホスト側が self.auth_btn / self.auth_status_label を持ち、
#       __init__ 内でウィジェット生成後に self._init_google_auth() を呼ぶこと。

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import QMessageBox

import google_auth_helper


class GoogleAuthMixin:
    # authorize_google_sheets()のコールバックは別スレッドから呼ばれるため、
    # シグナル経由でGUIスレッドに安全にディスパッチする
    auth_done = pyqtSignal(bool, str)

    def _init_google_auth(self):
        self.auth_done.connect(self._on_auth_done)

    def _authorize_google(self):
        self.auth_btn.setEnabled(False)
        self.auth_status_label.setText("🔄 ブラウザで認可中…")
        google_auth_helper.authorize_google_sheets(
            lambda ok, msg: self.auth_done.emit(ok, msg)
        )

    def _on_auth_done(self, ok, msg):
        if ok:
            QMessageBox.information(self, "Google 認証", msg)
        else:
            QMessageBox.warning(self, "Google 認証", msg)
        self._refresh_auth_label()
        if hasattr(self, "account_badge"):
            self.account_badge.refresh()

    def _refresh_auth_label(self):
        ok = google_auth_helper.has_token()
        self.auth_status_label.setText("✅ 認可済み" if ok else "❌ 未認可")
        self.auth_btn.setEnabled(not ok)
