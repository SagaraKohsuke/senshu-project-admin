# 泉州会館 食事申し込み管理システム

泉州会館の食事申し込みを管理するためのGoogle Apps Scriptベースの管理システムです。

## 🎯 主な機能

### 📅 予約管理
- **月間カレンダー表示**: 朝食・夕食の予約状況を一目で確認
- **4列レスポンシブデザイン**: 本日・明日の朝食・夕食予約状況を見やすく表示
- **リアルタイム更新**: 予約データの即座反映

### 🍽️ メニュー管理
- **朝食・夕食メニューの設定**: 日付別のメニュー管理
- **自由な変更**: 締切制限なしでいつでも変更可能

### 📊 食事原紙管理（自動化システム）
- **自動シート作成**: 毎月1日00:00に新しい月のシートを自動生成
- **自動データ更新**: 毎日12:00に当日以降の最新予約データを反映
- **ワンクリックアクセス**: 「食事原紙を確認」ボタンで即座にスプレッドシート表示

## 🏗️ システム構成

### データベース
- **メイン予約データ**: `17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk`
- **食事原紙**: `17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4`

### ファイル構成
```
📁 senshu-project-admin/
├── 📄 admin_index.html      # メイン管理画面（Vue.js 2.6.14）
├── 📄 admin_main.gs         # エントリーポイント・設定
├── 📄 admin_menu.gs         # メニュー管理機能
├── 📄 admin_submission.gs   # 食事原紙管理・自動化システム
├── 📄 admin_calendar.gs     # カレンダー表示機能
├── 📄 admin_utils.gs        # ユーティリティ関数
├── 📄 appsscript.json       # Apps Script設定
└── 📄 README.md             # このファイル
```

## 🚀 セットアップ手順

### 1. 初回設定
1. Google Apps Script エディタ (script.google.com) を開く
2. このプロジェクトをアップロード
3. 必要な権限を許可

### 2. 自動化トリガーの設定
```javascript
// admin_submission.gs で以下を実行
setupTriggers()
```

**設定されるトリガー:**
- 🕐 **毎月1日 00:00**: `createMonthlySheet()` - 新月度シート作成
- � **毎日 12:00**: `updateDailyMealSheet()` - 当日以降の予約データ更新

### 3. 初月シートの手動作成（必要に応じて）
現在の月のシートが存在しない場合：
```javascript
// admin_submission.gs で実行
createMonthlySheet()
```

## 💻 技術仕様

### フロントエンド
- **Vue.js 2.6.14**: リアクティブUIフレームワーク
- **CSS Grid**: レスポンシブ4列レイアウト
- **Android対応**: フォント表示最適化

### バックエンド
- **Google Apps Script**: サーバーレス実行環境
- **Google Spreadsheet API**: データ操作
- **Time-driven Triggers**: 自動化スケジューリング

### データ構造
```javascript
// 予約データ例
{
  date: "2025-08-15",
  breakfast: { users: [{userId: 1, name: "山田太郎"}], count: 1 },
  dinner: { users: [{userId: 2, name: "佐藤花子"}], count: 1 }
}
```

## 📋 運用フロー

### 日次運用
1. **12:00**: 自動でその日以降の予約データが食事原紙に反映
2. **管理者確認**: 「食事原紙を確認」ボタンで最新状況を確認

### 月次運用
1. **毎月1日 00:00**: 新しい月のシートが自動作成
2. **月末**: 前月データの最終確認

### 手動操作
- **メニュー変更**: 管理画面から随時可能
- **予約状況確認**: リアルタイムで4列表示
- **食事原紙確認**: ワンクリックでスプレッドシート表示

## 🔧 カスタマイズ

### スプレッドシートID変更
```javascript
// admin_main.gs
const spreadsheetId = "新しいメインデータベースID";
const mealSheetId = "新しい食事原紙ID";
```

### トリガー時間変更
```javascript
// admin_submission.gs の setupTriggers() 内
ScriptApp.newTrigger('updateDailyMealSheet')
  .timeBased()
  .everyDays(1)
  .atHour(新しい時間)  // 例: 19 (19:00)
  .create();
```

## 🐛 トラブルシューティング

### よくある問題

**Q: 食事原紙が更新されない**
A: トリガーが正しく設定されているか確認
```javascript
// 現在のトリガー確認
ScriptApp.getProjectTriggers().forEach(t => console.log(t.getHandlerFunction()));
```

**Q: 予約データが表示されない**
A: スプレッドシートの権限とIDを確認

**Q: Android端末で文字が見えない**
A: CSSで `!important` フラグ付きフォント設定済み

### ログ確認
```javascript
// Google Apps Script エディタで実行ログを確認
console.log("デバッグ情報");
```

## 📞 サポート

- **システム更新**: このREADMEと併せてコード内コメントを参照
- **データバックアップ**: Google Spreadsheetの版履歴機能を活用
- **権限管理**: Google Apps Scriptの共有設定で管理

## 📈 更新履歴

- **v2.0** (2025-08): 食事原紙自動化システム導入
- **v1.5** (2025-08): 4列レスポンシブデザイン実装
- **v1.4** (2025-08): 締切制限撤廃
- **v1.3** (2025-08): Android端末対応
- **v1.0** (2025): 初期リリース

---

**🏛️ 泉州会館 食事申し込み管理システム** - シンプル、自動化、効率的