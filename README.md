# NFC Equipment Manager

NFC リーダーを使った Python ベースの備品管理システムです（デスクトップ GUI 版）。  
SBC（Raspberry Pi 等）にモニター・キーボード・NFC リーダーを接続し、スタンドアローンで動作します。

> V2（Web ブラウザ版）はこちら: [NFC_EquipmentManager_V2](https://github.com/ReBot-dev/NFC_EquipmentManager_V2)

## 概要

社員証と物品に貼った NFC タグをスキャンするだけで、Google スプレッドシートに貸出・返却記録を残せます。  
社員証登録時に社内メールアドレスを登録することで、Slack Bot と連携して貸出通知・返却リマインダーを Slack DM で受け取れます。

## 機能

- NFC カードの貸出 / 返却登録（社員証 → 物品の順にタッチ、またはその逆）
- 未登録 NFC タグの社員証・物品としての新規登録
- 貸出中一覧 / 返却履歴 / 社員一覧 / 物品一覧の表示
- 不具合報告フォーム
- Slack Bot 連携（任意）
  - 社員証登録完了通知
  - 貸出完了通知
  - 返却期限超過リマインダー
  - `/help` `/list` スラッシュコマンド対応

## ハードウェア構成

| 機器 | 用途 |
|------|------|
| Raspberry Pi / Khadas VIM 等 SBC | 実行環境 |
| PC/SC 対応 NFC リーダー（例: Sony RC-S380） | NFC タグ読み取り |
| モニター + キーボード | GUI 操作 |

## Google スプレッドシートの準備

スプレッドシート名は `Equipment_Manager`（ソースコード内で変更可）。

以下のシートを作成してください：

| シート名 | 列構成 |
|----------|--------|
| 社員マスタ | 氏名 / IDm / メールアドレス / Slack ID |
| 物品マスタ | 物品名 / IDm / 貸出中の社員 / 最終貸出日時 |
| 貸出中一覧 | 申請日時 / 申請者 / 物品名 / 返却予定日 |
| 返却履歴 | 返却日時 / 物品名 / 返却者 / 予定返却日 |
| 不具合報告 | 対応状況 / 報告日時 / 報告者 / 不具合内容 |

Google Cloud Console でサービスアカウントを作成し、スプレッドシートへの編集権限を付与してから JSON キーファイルをダウンロードしてください。

## インストール

```bash
git clone https://github.com/ReBot-dev/NFC_EquipmentManager.git
cd NFC_EquipmentManager

pip install FreeSimpleGUI gspread pyscard
```

ダウンロードしたサービスアカウント JSON をプロジェクトフォルダに配置し、`Equipment_Manager_1.1.py` の以下の行を編集します：

```python
gc = gspread.service_account(filename=r"replace_with_your_service_account_json_file.json")
```

## 起動

```bash
python Equipment_Manager_1.1.py
```

## 操作方法

- **Tab キー** でボタン間を移動、**Space / Enter** で決定
- **Esc キー** でメインメニューに戻る
- フォーカス中のボタンは緑色にハイライトされます

## Slack Bot の準備（任意）

1. スプレッドシートのメニューから「拡張機能」→「Apps Script」を開く
2. `CheckIDandSendDM.js` と `command.js` の内容をスクリプトエディタに貼り付ける
3. スクリプトプロパティに `SLACK_BOT_TOKEN` を設定する
4. `onChangeHandler` を「スプレッドシートの変更時」トリガーに登録する
5. `remindUnreturnedItems` を時間主導型トリガー（例：毎日 18:00）に登録する
6. `/help` コマンド用に `command.js` の `doPost` を Web アプリとしてデプロイし、Slack のスラッシュコマンドに登録する

## 依存ライブラリ

| ライブラリ | 用途 |
|------------|------|
| FreeSimpleGUI | デスクトップ GUI フレームワーク |
| gspread | Google Sheets API クライアント |
| pyscard | PC/SC NFC リーダーインターフェース |
