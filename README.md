# diff_to_excel

2つのディレクトリ（またはファイル）間の差分を検出し、差分のビジュアル（HTML Diff）を画像化して、目次・リンク付きの洗練された Excel レポート（`.xlsx`）へ自動出力する Python スクリプトです。

---

## 📌 主な機能・特徴

- **ディレクトリ・ファイルの柔軟な差分比較**:
  - ルートディレクトリ全体の再帰的チェック
  - 特定のフォルダや単一ファイルのピンポイント指定比較
- **見やすいビジュアル差分**:
  - Python 標準の `difflib.HtmlDiff` で差分HTMLを生成し、見やすいスタイル（折り返し・余白調整）を適用
  - `wkhtmltoimage` と `Pillow` で差分を自動画像化・リサイズ最適化
- **デザインされた Excel レポート自動生成**:
  - **目次シート**: 差分があるファイル一覧と各差分シートへのハイパーリンクを自動生成
  - **差分シート**: ファイルごとに個別シートを作成し、差分画像を美しく配置（グリッド線非表示、ヘッダースタイル適用）
- **クリーンな実行**:
  - 一時生成される画像やHTMLファイルは処理後に自動削除
  - 実行サマリー（チェック対象、差分件数、ファイル一覧）をコンソールに出力

---

## 🛠 前提条件 (Prerequisites)

本ツールは HTML を画像へ変換するために **wkhtmltopdf (`wkhtmltoimage`)** を使用します。事前にインストールが必要です。

### 1. `wkhtmltoimage` のインストール

- **Windows**:
  - [wkhtmltopdf 公式ダウンロードページ](https://wkhtmltopdf.org/downloads.html) からインストーラーを入手してインストールしてください。
  - 通常は `C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe` にインストールされます。
- **macOS (Homebrew)**:
  ```bash
  brew install --cask wkhtmltopdf
  ```
  （実行パス例: `/usr/local/bin/wkhtmltoimage` または `/opt/homebrew/bin/wkhtmltoimage`）
- **Ubuntu / Debian**:
  ```bash
  sudo apt-get update && sudo apt-get install -y wkhtmltopdf
  ```

### 2. Python 依存パッケージのインストール

Python 3.8 以上が推奨されます。

```bash
pip install -r requirements.txt
```

#### `requirements.txt` の内容:
- `imgkit`: HTML to Image 変換ラッパー
- `openpyxl`: Excel ファイル操作
- `Pillow`: 画像リサイズ・最適化

---

## 🚀 使い方

### 1. スクリプト内の設定を編集

`diff_to_excel.py` の上部にある「設定エリア」をお使いの環境に合わせて編集します。

```python
# 設定エリア
# 比較元のベースとなるルートディレクトリ
DIR1_BASE = os.path.normpath(r"C:\path\to\compare1")
DIR2_BASE = os.path.normpath(r"C:\path\to\compare2")

# 比較したい「フォルダ」「ファイル」を相対パスのリストで指定
# 空リスト [] にすると、DIR1_BASE 配下を丸ごと再帰チェックします
TARGET_SUBDIRS = [
    # r"src/folder1",
    # r"src/file.txt",
]

# 出力するExcelファイル名
OUTPUT_EXCEL = "diff_report.xlsx"

# wkhtmltoimage の実行ファイルパス
# Windows の例:
WKHTMLTOIMAGE_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe"
# macOS / Linux の例:
# WKHTMLTOIMAGE_PATH = "/usr/local/bin/wkhtmltoimage"
```

### 2. 差分比較の実行

```bash
python diff_to_excel.py
```

実行すると、コンソールに進捗が出力され、完了時に `diff_report.xlsx` が生成されます。

---

## 📊 出力レポートの構成

1. **目次シート**:
   - 差分が検出されたファイルが一覧表示されます。
   - 「シートへ移動→」をクリックすると該当ファイルの差分シートにジャンプします。
2. **差分シート（ファイルごと）**:
   - シート名は対象ファイル名に基づいて設定されます。
   - 上部に対象ファイルパス、その下に色分けされた差分スクリーンショットが配置されます。

---

## 📁 ディレクトリ構成

```text
diff_to_excel/
├── diff_to_excel.py    # メインスクリプト
├── requirements.txt    # 必要なPythonライブラリ一覧
├── .gitignore          # 出力Excel等を除外
└── README.md           # 本ドキュメント
```

---

## 📝 注意事項

- ファイルの文字コードは `UTF-8` を前提として読み込みます（エラーはスキップ）。
- ファイル数やサイズが大きい場合、画像化処理に時間がかかることがあります。
