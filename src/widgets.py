#widgets.py
#
# 共通ウィジェット：メンション行入力

from PyQt6.QtWidgets import QWidget, QHBoxLayout, QLineEdit, QPushButton


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
