# Claude Code - Project Environment Guide

## Python 環境

このPCでは **Spyder 6.1.3** を使って開発しています。
Python は Spyder に同梱されたランタイム環境を統一して使用します。

### メイン Python 環境 (Spyder runtime)

| 項目 | 値 |
|------|-----|
| Pythonバージョン | **3.12.11** |
| 実行ファイル | `C:\ProgramData\spyder-6\envs\spyder-runtime\python.exe` |
| パッケージ管理 | pip |
| Spyderバージョン | 6.1.3 |

```bash
# このPCでPythonスクリプトを実行する場合
/c/ProgramData/spyder-6/envs/spyder-runtime/python.exe script.py
```

```bash
# パッケージをインストールする場合
/c/ProgramData/spyder-6/envs/spyder-runtime/Scripts/pip.exe install <package>
```

### インストール済み主要パッケージ

| パッケージ | バージョン | 用途 |
|-----------|-----------|------|
| numpy | 2.4.2 | 数値計算 |
| pandas | 3.0.0 | データ処理 |
| matplotlib | 3.10.8 | グラフ描画 |
| scipy | 1.17.0 | 科学技術計算 |
| openpyxl | 3.1.5 | Excel読み書き |
| requests | 2.32.5 | HTTP通信 |
| pywin32 (win32com) | 311 | Windows/Excel COM操作 |

### Spyder の設定

Spyder コンソール（IPython kernel）は spyder-runtime Python を使用するよう設定：

- `Tools` → `Preferences` → `Python interpreter`
- → "Use the following Python interpreter"
- → `C:\ProgramData\spyder-6\envs\spyder-runtime\python.exe`

### 注意: conda 環境について

`C:\Users\seijis\.conda\envs\data-analysis\` (Python 3.11.15) は旧 kernel 環境です。
Spyder の kernel を spyder-runtime に切り替えた後は不要となるため削除予定です。

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

Windows タスクスケジューラで毎日 **2:00** に自動実行するよう設定済みです。

```bash
# 手動で今すぐ保存したい場合
bash /c/Users/seijis/Claude_Code/daily_commit.sh
```

ログ: `C:\Users\seijis\Claude_Code\daily_commit.log`

## Claude Code 実行時の注意事項

1. **Pythonスクリプト実行**: `C:\ProgramData\spyder-6\envs\spyder-runtime\python.exe` を使用
2. **パッケージインストール**: `C:\ProgramData\spyder-6\envs\spyder-runtime\Scripts\pip.exe install <package>`
3. **Excel操作コード**: `win32com.client` / `openpyxl` はどちらも動作確認済み
4. **Spyder起動**: スタートメニューの「Spyder 6」から起動
