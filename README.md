# 研究室予約システム

山口大学 制御工学研究室向けの、Google Apps Script (GAS) および Google スプレッドシート、Google カレンダーを連携させた研究室利用予約システムです。
ほぼGeminiで作ったから適当。
開発者：2026年度制御研M2 内冨淳  メアド：junjun20091701@gmail.com



## ローカルフォルダ構成（LabRoomReservation/）

ローカル環境では、以下の構成でプログラムソースコードおよびドキュメント類をGit管理しています。

- **`src/`**：GASのソースコード一式（`.js`、`.html`、`appsscript.json` 等が格納されています。`clasp` の対象）
- **`docs/`**：設計書や仕様書などのドキュメント類（現在は空）
- **`html/`**：Web用の静的HTML等の保管場所（現在は空）
- **`tests/`**：テスト用コードの保管場所（現在は空）
- **`README.md`**：このシステム説明書（本ファイル）

---

## Google ドライブのフォルダ構成

Google ドライブ上は以下の構成でデータを分類・運用します。

- **`01_Development/`**：開発・検証環境（開発完了時の控え・テスト環境として利用）
  - `Spreadsheet/`：検証用のスプレッドシート（本番スプシのコピー等）
  - `AppScript/`：検証用のGASプロジェクトへのショートカット
  - `Calendar/`：検証用カレンダーに関する情報
- **`02_Production/`**：本格運用環境（本番環境）
  - `Spreadsheet/`：実際に稼働する「本番用スプレッドシート」（Reservations, Rules, Inquiries等）
  - `AppScript/`：本番GASプロジェクトへのショートカット、および本番WebアプリのURL
  - `Calendar/`：本番用カレンダーのID（`seigyo.reserve.system@gmail.com`）を記載したメモ
- **`03_Documents/`**：マニュアルやルール説明などの資料
- **`04_Backup/`**：過去の予約データや定期バックアップ
- **`05_GitExport/`**：ローカル（`LabRoomReservation`）からGit管理用に書き出された資産、あるいは連携用の中間データを一時的に扱うためのフォルダ

---

## 各スプレッドシート（シート）の構成

本番スプレッドシート（`02_Production/Spreadsheet` 内）は、以下のシートで構成されています。

1. **`Reservations`**：すべての予約データが蓄積されるメインシート。
2. **`Rules`**：特定の曜日、日付、期間、時間帯の予約制限を管理するルールシート。
3. **`Inquiries`**：利用者からの「利用不可日時の申請」「不具合報告」「要望」などが自動で蓄積されるシート。
4. **`ReservationHistory`**：新規作成、変更、キャンセルの操作ログが保存される履歴シート。
5. **`MailQueue`**：自動送信するメールの一時格納用シート。
6. **`SystemLog`**：エラーや実行情報を記録するシステムログシート。

---

## 開発・運用管理手順（claspコマンド）

ローカル環境でのコード管理やGASへの反映には `clasp` コマンドを使用します。PowerShell上で `LabRoomReservation` フォルダに移動して実行します。

### 1. GASから最新コードをローカルに取得する
```bash
clasp pull