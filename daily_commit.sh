#!/bin/bash
# daily_commit.sh
# 毎日の作業内容をGitHubに自動保存するスクリプト

REPO_DIR="C:/Users/seijis/Claude_Code"
DATE=$(date +"%Y-%m-%d")

cd "$REPO_DIR" || exit 1

# 変更がなければ終了
if git diff --quiet && git diff --staged --quiet && [ -z "$(git ls-files --others --exclude-standard)" ]; then
    echo "[$DATE] 変更なし - コミットをスキップします"
    exit 0
fi

# 全ファイルをステージング
git add -A

# 日付付きでコミット
git commit -m "Daily update: $DATE"

# GitHubにプッシュ
git push origin main

echo "[$DATE] GitHubへの保存が完了しました"
