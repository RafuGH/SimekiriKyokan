#theme.py
#
# カラートークンとスタイルシート（Fluent / Windowsライク）

class Colors:
    # ライト
    L_BG        = "#f3f3f3"   # ウィンドウ背景
    L_SURFACE   = "#ffffff"   # カード・入力背景
    L_SURFACE2  = "#f9f9f9"   # サブ背景
    L_BORDER    = "#e0e0e0"   # ボーダー
    L_TEXT      = "#1b1b1b"   # メインテキスト
    L_TEXT2     = "#616161"   # サブテキスト
    L_ACCENT    = "#0078d4"   # Fluent アクセントブルー
    L_ACCENT_H  = "#106ebe"   # ホバー
    L_ACCENT_P  = "#005a9e"   # プレス
    L_DANGER    = "#c50f1f"   # 削除・警告
    L_SUCCESS   = "#107c10"   # 成功
    L_HEADER    = "#fafafa"   # タイトルバー背景

    # ダーク
    D_BG        = "#202020"
    D_SURFACE   = "#2c2c2c"
    D_SURFACE2  = "#383838"
    D_BORDER    = "#404040"
    D_TEXT      = "#ffffff"
    D_TEXT2     = "#9d9d9d"
    D_ACCENT    = "#60cdff"
    D_ACCENT_H  = "#4ec9f0"
    D_ACCENT_P  = "#3ab8e0"
    D_DANGER    = "#ff99a4"
    D_SUCCESS   = "#6ccb5f"
    D_HEADER    = "#1c1c1c"


def make_stylesheet(dark: bool) -> str:
    c = Colors
    if dark:
        bg, surf, surf2, brd = c.D_BG, c.D_SURFACE, c.D_SURFACE2, c.D_BORDER
        txt, txt2            = c.D_TEXT, c.D_TEXT2
        acc, acc_h, acc_p    = c.D_ACCENT, c.D_ACCENT_H, c.D_ACCENT_P
        danger               = c.D_DANGER
        hdr                  = c.D_HEADER
    else:
        bg, surf, surf2, brd = c.L_BG, c.L_SURFACE, c.L_SURFACE2, c.L_BORDER
        txt, txt2            = c.L_TEXT, c.L_TEXT2
        acc, acc_h, acc_p    = c.L_ACCENT, c.L_ACCENT_H, c.L_ACCENT_P
        danger               = c.L_DANGER
        hdr                  = c.L_HEADER

    return f"""
/* ── ベース ── */
QWidget {{
    font-family: "Segoe UI", "Meiryo UI", "Yu Gothic UI", sans-serif;
    font-size: 13px;
    color: {txt};
    background-color: {bg};
}}
QScrollArea, QScrollArea > QWidget > QWidget {{ background: {bg}; border: none; }}

/* ── タイトルバー相当 ── */
QWidget#titlebar {{
    background-color: {hdr};
    border-bottom: 1px solid {brd};
}}

/* ── カードフレーム ── */
QFrame#card {{
    background-color: {surf};
    border: 1px solid {brd};
    border-radius: 8px;
}}

/* ── セクションヘッダー ── */
QLabel#section_hdr {{
    font-size: 11px;
    font-weight: 600;
    color: {txt2};
    letter-spacing: 0.8px;
    padding: 14px 0 2px 0;
}}

/* ── フィールドラベル ── */
QLabel#field_lbl {{
    font-size: 12px;
    color: {txt2};
    padding: 4px 0 1px 0;
}}

/* ── 説明ラベル ── */
QLabel#desc_lbl {{
    font-size: 11px;
    color: {txt2};
    padding: 0;
}}

/* ── 汎用ラベル ── */
QLabel {{ color: {txt}; background: transparent; }}

/* ── 入力フィールド共通 ── */
QLineEdit, QSpinBox, QComboBox, QTimeEdit, QDateEdit {{
    background: {surf};
    border: 1px solid {brd};
    border-radius: 6px;
    padding: 7px 12px;
    color: {txt};
    min-height: 32px;
    selection-background-color: {acc};
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus,
QTimeEdit:focus, QDateEdit:focus {{
    border: 1.5px solid {acc};
}}
QLineEdit::placeholder {{ color: {txt2}; }}
QComboBox::drop-down {{ border: none; width: 20px; }}
QComboBox QAbstractItemView {{
    background: {surf};
    border: 1px solid {brd};
    selection-background-color: {acc};
    color: {txt};
}}

/* ── プライマリボタン ── */
QPushButton#btn_primary {{
    background-color: {acc};
    color: {"#000000" if dark else "#ffffff"};
    border: none;
    border-radius: 6px;
    padding: 8px 22px;
    font-weight: 600;
    min-height: 32px;
}}
QPushButton#btn_primary:hover  {{ background-color: {acc_h}; }}
QPushButton#btn_primary:pressed {{ background-color: {acc_p}; }}
QPushButton#btn_primary:disabled {{ background-color: {brd}; color: {txt2}; }}

/* ── セカンダリボタン ── */
QPushButton {{
    background-color: {surf};
    color: {txt};
    border: 1px solid {brd};
    border-radius: 6px;
    padding: 7px 18px;
    min-height: 32px;
}}
QPushButton:hover  {{ background-color: {surf2}; border-color: {txt2}; }}
QPushButton:pressed {{ background-color: {brd}; }}
QPushButton:disabled {{ color: {txt2}; }}

/* ── 危険ボタン ── */
QPushButton#btn_danger {{
    color: {danger};
    border-color: {danger};
    background: transparent;
}}
QPushButton#btn_danger:hover {{ background-color: {"#3a1010" if dark else "#fff0f0"}; }}

/* ── ヘルプボタン ── */
QToolButton#btn_help {{
    background: transparent;
    border: 1px solid {brd};
    border-radius: 10px;
    color: {txt2};
    font-size: 11px;
    font-weight: 600;
    min-width: 20px;
    max-width: 20px;
    min-height: 20px;
    max-height: 20px;
    padding: 0;
}}
QToolButton#btn_help:hover {{ border-color: {acc}; color: {acc}; background: {surf2}; }}

/* ── チェックボックス ── */
QCheckBox {{ spacing: 8px; color: {txt}; }}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border: 1px solid {brd};
    border-radius: 3px;
    background: {surf};
}}
QCheckBox::indicator:checked {{
    background: {acc};
    border-color: {acc};
}}

/* ── テーブル ── */
QTableWidget {{
    background: {surf};
    border: 1px solid {brd};
    border-radius: 6px;
    gridline-color: {brd};
    outline: none;
}}
QTableWidget::item {{ padding: 6px 10px; }}
QTableWidget::item:selected {{
    background: {"#004f8c" if dark else "#cce4f7"};
    color: {txt};
}}
QTableWidget::item:alternate {{ background: {surf2}; }}
QHeaderView::section {{
    background: {surf2};
    border: none;
    border-bottom: 1px solid {brd};
    padding: 6px 10px;
    font-weight: 600;
    font-size: 11px;
    color: {txt2};
}}

/* ── スクロールバー ── */
QScrollBar:vertical {{
    width: 8px; background: transparent; margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {brd}; border-radius: 4px; min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {txt2}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}

/* ── テキストブラウザ（ヘルプウィンドウ） ── */
QTextBrowser {{
    background: {surf};
    border: none;
    color: {txt};
    font-size: 13px;
    line-height: 1.6;
}}

/* ── セパレータ ── */
QFrame[frameShape="4"] {{ color: {brd}; }}
"""


def apply_theme(widget, dark: bool = None):
    """
    保存されているテーマ設定（既定はライト）を widget に適用する。
    dark を明示した場合はその値を優先する。
    """
    import app_settings
    if dark is None:
        dark = app_settings.is_dark()
    widget.setStyleSheet(make_stylesheet(dark))
    return dark
