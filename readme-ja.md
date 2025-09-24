# AudioNormalizer

FFmpeg を使って音声のラウドネスを解析・正規化する Windows 向けデスクトップ GUI アプリです。PyQt5 製で、バックグラウンドスレッドにより UI を固まらせず、メタデータやアルバムアートを保持したまま、ターゲット LUFS、サンプリング周波数、ビットレートモード（VBR/CBR）、ビットレートを指定できます。

## システム要件（実行）

- Windows、macOS、またはLinux

## システム要件（開発）

- Windows、macOS、またはLinux
- Python 3.10以上
- PyQt5

## 開発について

### 依存関係のインストール

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 直接実行
```bash
python audio_normalizer.py
```

## 実行ファイルの作成

### Windows環境
PowerShellスクリプトを使用して実行ファイルを作成：
```powershell
.\make.ps1
```

### Unix系環境（Linux/macOS）
シェルスクリプトを使用して実行ファイルを作成：
```bash
bash ./make.sh
```

実行ファイルは `dist/` フォルダに生成されます。
