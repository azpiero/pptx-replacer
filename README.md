![Python](https://img.shields.io/badge/Python-3.7+-3776AB?logo=python&logoColor=white)
![python-pptx](https://img.shields.io/badge/python--pptx-1.0+-orange)
![Pillow](https://img.shields.io/badge/Pillow-10.0+-green)
![PyInstaller](https://img.shields.io/badge/PyInstaller-6.0+-yellow)
![License](https://img.shields.io/badge/License-MIT-blue)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

# PowerPoint画像一括置換システム 手順書

## 概要

このシステムは、複数のPowerPointファイル（.pptx）に含まれる特定の画像を、別の画像に一括で置き換えるためのPythonスクリプトです。

**主な機能:**
- 単一または複数のPPTXファイル内の画像を一括置換
- 画像の識別方法: MD5ハッシュ、ファイル名、ファイルサイズ
- 置換前の自動バックアップ
- PPTXファイル内の画像分析機能

---

## 必要環境

- Python 3.7以上
- 追加ライブラリ不要（標準ライブラリのみ使用）

---

## インストール

```bash
# スクリプトをダウンロード/コピー
curl -O https://your-server/replace_images.py

# 実行権限を付与（Linux/macOS）
chmod +x replace_images.py
```

---

## 基本的な使い方

### ステップ1: 置換対象の画像を分析する

まず、置換したい画像（現在PPTXに含まれている画像のオリジナル）を分析して、識別子を取得します。

```bash
python replace_images.py --analyze /path/to/target_image.png
```

**出力例:**
```
==================================================
画像分析結果
==================================================
ファイル名: logo.png
ファイルサイズ: 15234 bytes
MD5ハッシュ: a1b2c3d4e5f67890abcdef1234567890

使用例:
  --match-by hash --target a1b2c3d4e5f67890abcdef1234567890
  --match-by filename --target logo.png
  --match-by size --target 15234
==================================================
```

### ステップ2: PPTXファイル内の画像を確認する（オプション）

特定のPPTXファイルに含まれる画像を確認できます。

```bash
# 単一ファイルをスキャン
python replace_images.py --scan /path/to/presentation.pptx

# ディレクトリ内の全PPTXを分析
python replace_images.py --scan-dir /path/to/pptx_folder
```

**出力例:**
```
presentation.pptx 内の画像一覧:
--------------------------------------------------------------------------------
No.  ファイル名                 サイズ       MD5ハッシュ
--------------------------------------------------------------------------------
1    image1.png                 45678        a1b2c3d4e5f67890abcdef1234567890
2    logo.png                   15234        f0e1d2c3b4a59687abcdef1234567890
3    chart.png                  28901        1234567890abcdef1234567890abcdef
--------------------------------------------------------------------------------
合計: 3個の画像
```

### ステップ3: 画像を一括置換する

識別子と置換用画像を指定して、一括置換を実行します。

```bash
# ハッシュ値で置換（最も正確）
python replace_images.py \
    --directory /path/to/pptx_folder \
    --target a1b2c3d4e5f67890abcdef1234567890 \
    --replacement /path/to/new_image.png

# ファイル名で置換
python replace_images.py \
    --directory /path/to/pptx_folder \
    --match-by filename \
    --target logo.png \
    --replacement /path/to/new_logo.png

# ファイルサイズで置換
python replace_images.py \
    --directory /path/to/pptx_folder \
    --match-by size \
    --target 15234 \
    --replacement /path/to/new_image.png
```

---

## オプション一覧

| オプション | 短縮形 | 説明 |
|-----------|-------|------|
| `--analyze IMAGE` | | 画像ファイルを分析して識別子を表示 |
| `--scan PPTX` | | PPTXファイル内の画像一覧を表示 |
| `--scan-dir DIR` | | ディレクトリ内の全PPTXの画像を分析 |
| `--directory DIR` | `-d` | 置換対象のPPTXファイルがあるディレクトリ |
| `--target ID` | `-t` | 置換対象の識別子（ハッシュ、ファイル名、またはサイズ） |
| `--replacement PATH` | `-r` | 置換用画像のパス |
| `--match-by METHOD` | `-m` | マッチング方法: `hash`（デフォルト）, `filename`, `size` |
| `--output-dir DIR` | `-o` | 出力先ディレクトリ（省略時は元ファイルを上書き） |
| `--no-backup` | | バックアップを作成しない |
| `--no-recursive` | | サブディレクトリを検索しない |

---

## 使用例

### 例1: 会社ロゴの一括更新

全てのプレゼンテーションで古いロゴを新しいロゴに置き換える場合:

```bash
# 1. 古いロゴの識別子を取得
python replace_images.py --analyze old_logo.png

# 2. 一括置換を実行（出力先を別ディレクトリに）
python replace_images.py \
    --directory ./presentations \
    --match-by hash \
    --target abc123def456789... \
    --replacement ./new_logo.png \
    --output-dir ./updated_presentations
```

### 例2: 特定のファイル名の画像を置換

```bash
# background.pngという名前の画像を全て置換
python replace_images.py \
    --directory ./presentations \
    --match-by filename \
    --target background.png \
    --replacement ./new_background.png
```

### 例3: バックアップなしで直接上書き

```bash
python replace_images.py \
    --directory ./presentations \
    --target abc123def456789... \
    --replacement ./new_image.png \
    --no-backup
```

### 例4: サブディレクトリを除外して置換

```bash
python replace_images.py \
    --directory ./presentations \
    --target abc123def456789... \
    --replacement ./new_image.png \
    --no-recursive
```

---

## マッチング方法の選び方

| 方法 | 使用場面 | メリット | デメリット |
|-----|---------|---------|-----------|
| `hash` | 同一画像を確実に特定したい場合 | 最も正確、誤置換のリスクが低い | オリジナル画像が必要 |
| `filename` | PPTX内のファイル名が分かっている場合 | 簡単、直感的 | 同名の別画像も置換される可能性 |
| `size` | ハッシュやファイル名が不明な場合 | オリジナル不要 | 同サイズの別画像も置換される可能性 |

**推奨**: 可能な限り `hash` を使用してください。

---

## 出力結果の見方

```
見つかったPPTXファイル: 5件
--------------------------------------------------
✓ presentation1.pptx: 成功: 2個の画像を置換しました
✓ presentation2.pptx: 成功: 1個の画像を置換しました
○ presentation3.pptx: マッチする画像が見つかりませんでした
✓ presentation4.pptx: 成功: 3個の画像を置換しました
✗ presentation5.pptx: ファイルが見つかりません: ...

==================================================
処理完了
==================================================
処理ファイル数: 5
置換成功: 3ファイル（計6画像）
マッチなし: 1ファイル
失敗: 1ファイル
```

| 記号 | 意味 |
|-----|------|
| ✓ | 画像の置換に成功 |
| ○ | 処理成功だがマッチする画像なし |
| ✗ | エラー発生 |

---

## トラブルシューティング

### Q: 「マッチする画像が見つかりません」と表示される

**考えられる原因:**
1. 識別子が正しくない
2. PPTXファイル内に該当画像が存在しない
3. マッチング方法が適切でない

**対処法:**
```bash
# PPTXファイル内の画像を確認
python replace_images.py --scan /path/to/file.pptx

# 表示された情報と指定した識別子を比較
```

### Q: 置換後にPPTXが開けない

**対処法:**
1. バックアップファイル（.pptx.backup）から復元
2. 置換用画像のフォーマットを確認（元画像と同じ形式推奨）

### Q: 一部のファイルだけ置換されない

**考えられる原因:**
- 対象ファイルが別のプログラムで開かれている
- ファイルのアクセス権限がない
- 一時ファイル（~$で始まるファイル）は自動的に除外される

---

## 注意事項

1. **バックアップ**: デフォルトでバックアップが作成されますが、重要なファイルは事前に手動でもバックアップしてください

2. **画像フォーマット**: 置換用画像は、可能な限り元画像と同じフォーマット（PNG→PNG、JPG→JPG）を使用してください

3. **サイズの違い**: 置換用画像のサイズが元画像と異なる場合、スライド上での表示サイズは変わりません（画像が引き伸ばされる/縮小される可能性があります）

4. **テスト実行**: 本番実行前に、テスト用のPPTXファイルで動作確認することを推奨します

---

## Windows EXEファイルの作成

PyInstallerを使用してスタンドアロンのEXEファイルを作成できます。

### クイックスタート

```cmd
# 1. 仮想環境の作成（推奨）
python -m venv venv_build
venv_build\Scripts\activate

# 2. 依存関係のインストール
pip install Pillow python-pptx pyinstaller

# 3. ビルド実行
pyinstaller pptx_replacer.spec
```

ビルド成功後、`dist/PPTX画像置換ツール.exe` が作成されます。

詳細な手順やトラブルシューティングは [BUILD_WINDOWS_EXE.md](BUILD_WINDOWS_EXE.md) を参照してください。

---

## ライセンス

MIT License

---

## 更新履歴

- v1.0.0 (2025-01-XX): 初版リリース
