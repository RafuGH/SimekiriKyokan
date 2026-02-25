# 締切教官 (SimekiriKyokan)
Excelと連携する Windows用締切通知アプリ  
(Python / PyQt6 / Inno Setup)

## 📌 概要
締切教官は、Excelで管理している作業の締切をDiscordで自動通知するデスクトップアプリです。

・シンプルなGUI操作 
・締切〇日前に通知 
・Windowsタスクスケジューラ対応

## 🖥 動作環境
・Windows 10 / 11 (64bit)
・Excel (xlsx対応)
・管理者権限での初回インストール

## 📦 インストール方法（一般ユーザー向け）
1. Releases から最新版の `SimekiriKyokan_Setup_x.x.exe` をダウンロード  
2. セットアップを実行  
3. 指示に従ってインストール  

インストール後、スタートメニューから起動できます。

## 📊 使い方（簡潔）
1. 専用のExcelを生成
2. 生成したエクセルを参照
3. 教官の設定（何日前に通知するか）などの設定  
4. 新規作成ボタンを押すとスケジュールが登録されます  
詳しい説明はアプリ画面右上にあるボタンを押してマニュアルをご覧ください。

## 📁 プロジェクト構成

```
SimekiriKyokan/
│
├─ assets/                     # Excelテンプレート・マニュアル
│   ├─ SimekiriKyokan_Manual.pdf
│   ├─ Tasks.xlsx
│   └─ icon.ico
│
├─ installer/                # Inno Setup スクリプト
│   └─ SimekiriKyokan.iss
│
├─ src/                      # Pythonソースコード
│   ├─ simekiri_gui.py
│   └─ simekiri_notify.py
│
├─ README.md
└─ requirements.txt
```

## 使用技術
・Python 3.11
・PyQt6
・openpyxl
・pandas
・Windows Task Scheduler
・Inno Setup

## 🛠 開発環境構築（開発者向け）
```bash
git clone https://github.com/あなたのID/SimekiriKyokan.git
cd SimekiriKyokan
pip install -r requirements.txt
python simekiri_gui.py

pyinstaller src/simekiri_gui.py --onefile --clean --icon=data/icon.ico
```

## 🔐 Security

VirusTotal scan result:
[https://www.virustotal.com/gui/file/xxxxxxxxxxxxxxxx](https://www.virustotal.com/gui/file/98d7a1721b0a6552bf4bd5c0ee402948a48be05325c39769f12af2c8462c5897/detection)
