# GROUP フォルダ - 月別集計・グラフ生成システム

## 概要
予約受付データの月別集計とグラフ生成を行うPythonスクリプト群です。

## 主要ファイル

### 月別集計・グラフ生成
- **`monthly_line_chart.py`** - ローカル実行用メインスクリプト
  - 予約データの月別集計とグラフ生成
  - HP経由・その他経由で分けて処理
  - matplotlib + win32com.client使用

- **`AWS_monthky_line_chart.py`** - AWS Lambda用
  - openpyxl使用のExcelグラフ生成
  - S3連携機能付き

- **`monthly_consolidator.py`** - 月別データ統合
  - 複数月のデータを統合
  - Excelファイルの結合処理

### その他の分析・グラフ生成
- **`menu_pie.py`** - メニュー別集計（円グラフ）
- **`site_pie.py`** - サイト別集計（円グラフ）
- **`monthly_6_usage.py`** - 6ヶ月使用量分析
- **`AWS_monthly_6_usage.py`** - AWS Lambda版6ヶ月分析

### ユーティリティ
- **`excel_chart_to_image_v2.py`** - Excelグラフを画像に変換
- **`excel_to_powerpoint_simple.py`** - ExcelからPowerPoint変換
- **`test_excel.py`** - Excel操作テスト
- **`group_local.py`** / **`group_local_v2.py`** - ローカル実行用統合スクリプト

## 使用方法

### ローカル実行
```bash
# 月別集計グラフ生成
python monthly_line_chart.py

# データ統合
python monthly_consolidator.py

# その他分析
python menu_pie.py
python site_pie.py
```

### AWS Lambda実行
- Step Functions経由で実行
- 環境変数設定が必要（S3_BUCKET_NAME等）

## データフロー
```
CSVデータ → 月別集計 → グラフ生成 → Excel出力 → S3アップロード
```

このREADMEは、GROUPフォルダ内の各ファイルの役割を簡潔に説明し、使用方法とデータフローを分かりやすく示しています。

## 出力ファイル
- **月別集計_HP.xlsx** - HP経由データ
- **月別集計_その他.xlsx** - その他経由データ
- **results/** - 生成されたExcelファイル格納

## 必要なライブラリ
```bash
pip install -r requirements.txt
```

## 注意事項
- Windows環境では`win32com.client`が必要
- AWS Lambda環境では`xlsxwriter`レイヤーが必要
- 日本語フォント設定が必要（matplotlib使用時）

## 推奨されるファイル構成

以下のファイルをGitに含めることを推奨します：

### ✅ 含めるファイル
- **Pythonスクリプト**: `*.py`
- **設定ファイル**: `requirements.txt`, `README.md`
- **テストデータ**: `test_event.json`
- **サンプルデータ**: `results/`内の一部（必要に応じて）

### ❌ 除外するファイル
- **生成されたExcelファイル**: `*.xlsx`（results/内）
- **一時ファイル**: `temp_charts/`, `~$*`
- **キャッシュファイル**: `__pycache__/`, `*.pyc`
- **環境固有ファイル**: `.env`, `.vscode/`

## セキュリティ考慮事項

### 機密情報の除外
```gitignore
# 機密情報
*.key
*.pem
secrets/
credentials/
config.json
```

### AWS認証情報
- AWS認証情報は環境変数またはAWS CLI設定を使用
- ハードコードされた認証情報は絶対に含めない

## コミットメ