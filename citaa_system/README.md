# CITAA System - 千葉工業大学体育会本部 統合管理システム

Flet デスクトップアプリケーション

## 機能一覧

### 書記部
- 施設利用記録の管理
- 複数行入力対応
- タイムライン表示
- Excelカレンダー出力

### 財務部
- 収支管理
- 立替精算機能（未返金/返金済）
- 手許現金・通帳残高の自動計算
- 収支一覧表示

### 総務部
- 曜日担当設定（前期/後期/夏季休業/冬季休業）
- 出欠記録
- 出欠割合統計

### 渉外部
- スキャン・データ保存（Nodeサーバー連携）
- 課外活動記録（課外活動届/活動報告書）
- 宿泊・合宿リスト抽出

### 編集部
- DB編集（必要項目/パスワード/顧問情報）
- 各URL設定
- 名簿校正プログラム（PDF一括生成）

### イベント管理
- フォーム集計
- フォルダ一括作成（サブフォルダ対応）

### 管理&設定
- 団体・区分管理
- 本部員情報管理（CSV/Excel一括インポート）
- システム設定
- パスワード変更

## セットアップ

### 1. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 2. Google API認証設定

1. Google Cloud Consoleでプロジェクトを作成
2. Google Sheets API と Google Drive API を有効化
3. OAuth 2.0クライアントIDを作成
4. 認証情報JSONをダウンロードし、`auth/credentials.json`として保存

### 3. アプリケーション起動

```bash
python main.py
```

## 初回ログイン手順

1. 「Googleアカウントでログイン」をクリック
2. ブラウザでアカウントを選択
3. 「このアプリはGoogleで確認されていません」→「続行」をクリック
4. アクセス権限を許可して「続行」
5. 「認証完了」画面が表示されたら完了

## ファイル構成

```
citaa_system/
├── main.py              # メインアプリケーション
├── config.py            # 設定・定数
├── requirements.txt     # 依存関係
├── auth/
│   ├── __init__.py
│   ├── google_auth.py   # Google認証管理
│   └── credentials.json # Google API認証情報（要作成）
├── services/
│   ├── __init__.py
│   ├── sheets_service.py    # スプレッドシート操作
│   └── error_logger.py      # エラーログ
├── assets/              # アセットファイル
├── logs/               # ログファイル
└── .cache/             # キャッシュ
```

## 必要なスプレッドシートのシート構成

- Settings
- Clubs (ClubName, Category, Color, DisplayName)
- Members (StudentID, Name, Department, Role)
- Facilities (FacilityID, FacilityName)
- SecretaryLog (Date, Facility, ClubName, StartTime, EndTime, Note, CreatedAt)
- Finance (Date, Subject, Description, PaymentMethod, Income, Expense, ReimbursementStatus, CreatedAt)
- Attendance (Date, MemberName, Status, Period, CreatedAt)
- WeekdayAssign (Period, Weekday, Members)
- ExternalLog
- Bookmarks (Name, URL)
- RequiredItems (ClubName, StudentPhone, GuarantorPhone, Address)
- Passwords (ClubName, Password)
- Advisors (ClubName, Director, Advisor, Coach, CoachSub)
- Categories (CategoryName, Order)
- StudyWeeks (Period, StartDate, EndDate)

## 注意事項

- PDF生成機能を使用するには `reportlab` と `PyPDF2` が必要です
- Excel出力機能を使用するには `openpyxl` が必要です
- スキャナー連携にはNodeサーバーが必要です

## バージョン

- v1.0.0 (2026-02-26)
