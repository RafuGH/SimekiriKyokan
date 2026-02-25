# 締切教官 (SimekiriKyokan)
Excelと連携する Windows用締切通知アプリ  
(Python / PyQt6 / Inno Setup)

## 📌 概要

締切教官は、Excelで管理しているタスクの締切を自動通知するデスクトップアプリです。
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
1. GUIから専用のExcelを生成
2. Excelに締切日を入力  
3. 通知設定（何日前に通知するか）を設定  
4. 登録ボタンを押すとスケジュールが設定されます  

詳しい説明はアプリを起動して右上にあるHelpボタンをクリックしてマニュアルをご覧ください。

## 🗂 プロジェクト構成
SimekiriKyokan/
│
├─ src/ # Pythonソースコード
│ ├─ simekiri_gui.py
│ ├─ simekiri_notify.py
│ └─ simekiri_register.py
│
├─ data/ # Excelテンプレート・マニュアル
│ ├─ Tasks.xlsx
│ ├─ icon.ico
│ └─ SimekiriKyokan_Manual.pdf
│
├─ installer/ # Inno Setup スクリプト
│ └─ setup.iss
│
├─ requirements.txt
└─ README.md

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
