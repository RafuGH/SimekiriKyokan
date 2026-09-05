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
│   ├─ simekiri_gui.py        # メインGUI（エントリポイント）
│   ├─ simekiri_notify.py     # 通知処理のオーケストレーション
│   ├─ task_scheduler.py      # タスクスケジューラ／管理者権限まわり
│   ├─ task_manager_window.py # タスク管理ウィンドウ
│   ├─ task_edit_dialog.py    # タスク編集ダイアログ
│   ├─ widgets.py             # 共通ウィジェット（メンション行入力）
│   ├─ theme.py               # カラートークン／スタイルシート
│   ├─ help_widgets.py        # フィールドごとのコンテキストヘルプ
│   ├─ account_badge.py       # タイトルバーのGoogleアカウント表示
│   ├─ google_auth_mixin.py   # Google認可ボタンの共通ロジック
│   ├─ webhook_client.py      # Webhook送信（Discord/Slack/Teams等）
│   ├─ task_image.py          # 作業リスト画像の生成
│   ├─ data_loader.py         # Excel/Googleスプレッドシート読み込み
│   ├─ google_auth_helper.py  # Google OAuth連携
│   ├─ app_settings.py        # アプリ設定（テーマ等）の保存
│   ├─ submission_watcher.py  # 提出フォルダ監視→通知
│   ├─ update_checker.py      # 起動時のアップデート確認
│   └─ simekiri_2_1_run.bat   # 開発時のローカル起動用スクリプト
│
├─ README.md
└─ requirements.txt
```

## ✨ 主な機能

### 締切通知
Excel / Google スプレッドシートの作業リストを読み取り、締切が近い作業を担当者ごとに
Discord / Slack / Teams / Chatwork / Google Chat へ通知します（画像付き）。

### 確認待ち通知
進捗が「確認待ち」の作業を、担当者ではなくレビュアーに通知します。

### 提出フォルダ監視
OneDrive / Teams の同期フォルダなどを指定しておくと、そこに新しいファイルが
提出された・更新されたときに通知します。

- チェックのタイミングは、自動連絡のスケジュール実行時と「今すぐ実行」時です
- 有効化直後の初回スキャンでは通知しません（既存ファイルの全件通知を防ぐため）
- PCの起動状態に関わらず即時通知したい場合は、Power Automate の
  「ファイルが作成されたとき」→「HTTP」で Webhook に POST するフローと併用できます

### ライト / ダークテーマ
既定はライトモードです。画面右上のボタンでいつでも切り替えでき、設定は保存されます。

### アップデート告知
起動時に GitHub の最新リリースを確認し、新しいバージョンがあればお知らせします。
（通知は同じバージョンにつき1回だけです。ネットワークに繋がらない場合は何もしません）

## 🚀 リリース手順（配布側）

アップデート告知を機能させるには、次の手順でリリースしてください。

1. `src/update_checker.py` の `APP_VERSION` を新しい番号に更新する（例: `"2.2"`）
2. `installer/SimekiriKyokan.iss` の `AppVersion` と `OutputBaseFilename` も合わせる
3. PyInstaller と Inno Setup でセットアップ exe をビルドする
4. GitHub の Releases で **タグ `vX.Y`**（`APP_VERSION` と対応する形）を付けて公開し、
   セットアップ exe を添付する

利用者が次に締切教官を起動したとき、新しいバージョンが検知され、
ダウンロードページへの案内が表示されます。

## 今後の予定
- Googleスプレッドシート対応
- Teams, Slack対応
- 確認待ちの締切通知を受け取る人の指定

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
