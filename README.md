# NFC Equipment Manager

NFC リーダーを使った Python ベースの備品管理システムです（デスクトップ GUI 版）。  
SBC（Raspberry Pi 等）にモニター・キーボード・NFC リーダーを接続し、スタンドアローンで動作します。

A Python-based equipment management system using an NFC reader (desktop GUI version).  
Runs standalone on an SBC (e.g. Raspberry Pi) connected to a monitor, keyboard, and NFC reader.

> V2（Web ブラウザ版 / Web browser version）: [NFC_EquipmentManager_V2](https://github.com/ReBot-dev/NFC_EquipmentManager_V2)

---

## 概要 / Overview

社員証と物品に貼った NFC タグをスキャンするだけで、Google スプレッドシートに貸出・返却記録を残せます。  
社員証登録時に社内メールアドレスを登録することで、Slack Bot と連携して貸出通知・返却リマインダーを Slack DM で受け取れます。

Scan NFC tags on employee ID cards and items to record borrowing and returns in Google Sheets.  
By registering a company email address at employee card registration, the system integrates with Slack Bot to send borrow notifications and return reminders via Slack DM.

---

## 機能 / Features

- NFC カードの貸出 / 返却登録（社員証 → 物品の順にタッチ、またはその逆）
- 未登録 NFC タグの社員証・物品としての新規登録
- 貸出中一覧 / 返却履歴 / 社員一覧 / 物品一覧の表示
- 不具合報告フォーム
- Slack Bot 連携（任意）：登録完了通知・貸出通知・返却期限リマインダー・`/help` `/list` コマンド対応

---

- Borrow / return registration via NFC (touch employee card → item, or vice versa)
- New registration of unregistered NFC tags as employee cards or items
- View current borrow list / return history / employee list / item list
- Bug report form
- Optional Slack Bot integration: registration notifications, borrow notifications, overdue reminders, `/help` `/list` slash commands

---

## ハードウェア構成 / Hardware

| 機器 / Device | 用途 / Purpose |
|---|---|
| Raspberry Pi / Khadas VIM 等 SBC | 実行環境 / Runtime |
| PC/SC 対応 NFC リーダー（例: Sony RC-S380） | NFC タグ読み取り / NFC tag reading |
| モニター + キーボード | GUI 操作 / GUI input |

---

## Google スプレッドシートの準備 / Google Sheets Setup

スプレッドシート名は `Equipment_Manager`（ソースコード内で変更可）。  
Spreadsheet name: `Equipment_Manager` (configurable in source code).

以下のシートを作成してください / Create the following sheets:

| シート名 / Sheet | 列構成 / Columns |
|---|---|
| 社員マスタ | 氏名 / IDm / メールアドレス / Slack ID |
| 物品マスタ | 物品名 / IDm / 貸出中の社員 / 最終貸出日時 |
| 貸出中一覧 | 申請日時 / 申請者 / 物品名 / 返却予定日 |
| 返却履歴 | 返却日時 / 物品名 / 返却者 / 予定返却日 |
| 不具合報告 | 対応状況 / 報告日時 / 報告者 / 不具合内容 |

Google Cloud Console でサービスアカウントを作成し、スプレッドシートへの編集権限を付与してから JSON キーファイルをダウンロードしてください。  
Create a service account on Google Cloud Console, grant it edit access to the spreadsheet, and download the JSON key file.

---

## インストール / Installation

```bash
git clone https://github.com/ReBot-dev/NFC_EquipmentManager.git
cd NFC_EquipmentManager

pip install FreeSimpleGUI gspread pyscard
```

ダウンロードしたサービスアカウント JSON をプロジェクトフォルダに配置し、`Equipment_Manager_1.1.py` の以下の行を編集します：  
Place the service account JSON in the project folder and edit the following line in `Equipment_Manager_1.1.py`:

```python
gc = gspread.service_account(filename=r"replace_with_your_service_account_json_file.json")
```

---

## 起動 / Usage

```bash
python Equipment_Manager_1.1.py
```

**操作方法 / Controls:**
- **Tab** でボタン間を移動 / Move between buttons
- **Space / Enter** で決定 / Confirm
- **Esc** でメインメニューに戻る / Return to main menu
- フォーカス中のボタンは緑色にハイライト / Focused button is highlighted in green

---

## Slack Bot の準備 / Slack Bot Setup（任意 / Optional）

1. スプレッドシートのメニューから「拡張機能」→「Apps Script」を開く  
   Open "Extensions" → "Apps Script" from the spreadsheet menu
2. `CheckIDandSendDM.js` と `command.js` の内容を貼り付ける  
   Paste the contents of `CheckIDandSendDM.js` and `command.js`
3. スクリプトプロパティに `SLACK_BOT_TOKEN` を設定する  
   Set `SLACK_BOT_TOKEN` in script properties
4. `onChangeHandler` を「スプレッドシートの変更時」トリガーに登録する  
   Register `onChangeHandler` as an "on spreadsheet change" trigger
5. `remindUnreturnedItems` を時間主導型トリガー（例：毎日 18:00）に登録する  
   Register `remindUnreturnedItems` as a time-based trigger (e.g. daily at 18:00)
6. `command.js` の `doPost` を Web アプリとしてデプロイし、Slack スラッシュコマンドに登録する  
   Deploy `doPost` in `command.js` as a Web App and register it as a Slack slash command

---

## 依存ライブラリ / Dependencies

| ライブラリ / Library | 用途 / Purpose |
|---|---|
| FreeSimpleGUI | デスクトップ GUI / Desktop GUI framework |
| gspread | Google Sheets API クライアント / client |
| pyscard | PC/SC NFC リーダー / reader interface |

---

## ライセンス / License

MIT License
