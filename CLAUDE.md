# Claude Code - Project Environment Guide

## Python 環境

このPCには複数のPython環境があります。Claude Code でコードを実行する際は以下を参照してください。

### メイン開発環境 (Conda)

| 項目 | 値 |
|------|-----|
| 環境名 | `data-analysis` |
| Pythonバージョン | 3.11.15 (Anaconda) |
| 実行ファイル | `C:\Users\seijis\.conda\envs\data-analysis\python.exe` |
| パッケージ管理 | conda / pip |

```bash
# このPCでPythonスクリプトを実行する場合
/c/Users/seijis/.conda/envs/data-analysis/python.exe script.py

# または Windows フルパス
C:\Users\seijis\.conda\envs\data-analysis\python.exe script.py
```

### Spyder IDE

| 項目 | 値 |
|------|-----|
| Spyderバージョン | 6.1.3 |
| インストール先 | `C:\ProgramData\spyder-6\` |
| Spyder内部Python | 3.12.11 (Spyder runtime専用) |
| 実行ファイル | `C:\ProgramData\spyder-6\envs\spyder-runtime\Scripts\spyder.exe` |

> **注意**: Spyder の「Run」ボタンで実行されるコードは `data-analysis` conda env (Python 3.11.15) を使用します。
> Spyder 内部ランタイムの Python 3.12.11 は Spyder 自身の動作用であり、ユーザーコードの実行には使用しません。

### Python バージョン早見表

| 用途 | バージョン | 実行ファイル |
|------|-----------|-------------|
| 通常の開発・スクリプト実行 | **Python 3.11.15** | `C:\Users\seijis\.conda\envs\data-analysis\python.exe` |
| Spyder IDE 本体の動作 | Python 3.12.11 | `C:\ProgramData\spyder-6\envs\spyder-runtime\python.exe` |

## GitHub リポジトリ

| 項目 | 値 |
|------|-----|
| リモートURL | `https://github.com/seijis12345/claude-code-projects.git` |
| デフォルトブランチ | `main` |
| Gitユーザー | seijis12345 |

## ディレクトリ構成

```
C:\Users\seijis\Claude_Code\   ← 作業ディレクトリ (このリポジトリ)
```

## 毎日のGitHub保存

毎日の作業終了時に `daily_commit.sh` を実行することで、その日作成・変更したファイルをGitHubに自動保存します。

```bash
bash /c/Users/seijis/Claude_Code/daily_commit.sh
```

Windows タスクスケジューラで毎日 18:00 に自動実行するよう設定済みです。

## Claude Code 実行時の注意事項

1. **Pythonスクリプト実行**: 必ず `data-analysis` conda env の Python を使用
2. **パッケージインストール**: `C:\Users\seijis\.conda\envs\data-analysis\Scripts\pip.exe install <package>`
3. **condaコマンド**: `conda activate data-analysis` してから使用
4. **Spyder起動**: スタートメニューの「Spyder 6」から起動（直接パス実行不要）
