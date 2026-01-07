# Windows EXEファイル作成手順

## 必要なもの
- Windows PC
- Python 3.8以上

## 手順

### 1. 必要ファイルをWindowsに転送
以下のファイルをWindows PCに転送してください：
- `pptx_image_replacer_gui.py`
- `pptx_replacer.spec`

### 2. Python仮想環境の作成（推奨）
```cmd
cd (ファイルを配置したフォルダ)
python -m venv venv_build
venv_build\Scripts\activate
```

### 3. 必要ライブラリのインストール
```cmd
pip install Pillow python-pptx pyinstaller
```

### 4. EXEファイルのビルド

#### 方法1: specファイルを使用（推奨）
```cmd
pyinstaller pptx_replacer.spec
```

#### 方法2: コマンドラインオプションで直接ビルド
```cmd
pyinstaller --onefile --windowed --name "PPTX画像置換ツール" pptx_image_replacer_gui.py
```

### 5. 出力ファイルの確認
ビルドが成功すると、`dist`フォルダ内に実行ファイルが作成されます：
```
dist/
  └── PPTX画像置換ツール.exe
```

### 6. 動作確認
1. `dist`フォルダ内の`.exe`ファイルをダブルクリック
2. GUIが起動することを確認
3. 画像の選択と置換が正常に動作することを確認

## トラブルシューティング

### エラー: tkinter not found
Pythonインストール時にtk/tclをインストールしてください：
```cmd
# Pythonを再インストールし、"tcl/tk and IDLE"にチェックを入れる
```

### エラー: PIL module not found
```cmd
pip install --upgrade Pillow
pyinstaller --clean pptx_replacer.spec
```

### EXEファイルサイズを小さくしたい
`pptx_replacer.spec`を編集：
```python
upx=True,  # UPX圧縮を有効化（既に有効）
```

さらに圧縮したい場合：
```cmd
pip install upx-windows
pyinstaller --clean pptx_replacer.spec
```

### コンソール出力を確認したい（デバッグ用）
`pptx_replacer.spec`を編集：
```python
console=True,  # Falseから変更
```

## 配布方法

### 単体配布
`dist/PPTX画像置換ツール.exe` を配布

### インストーラー作成（オプション）
Inno SetupやNSISを使用してインストーラーを作成可能です。

## 注意事項
- 初回起動時、Windowsセキュリティの警告が表示される場合があります
- ウイルス対策ソフトが誤検知する場合があります（PyInstallerの既知の問題）
- 配布前に必ずウイルススキャンを実行してください
