# Windows タスクスケジューラに毎日18:00の自動コミットタスクを登録

$TaskName = "ClaudeCode_DailyGitCommit"
$GitBash = "C:\Program Files\Git\bin\bash.exe"
$Script = "C:\Users\seijis\Claude_Code\daily_commit.sh"
$LogFile = "C:\Users\seijis\Claude_Code\daily_commit.log"

# 既存タスクを削除（再登録のため）
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

# アクション定義
$Action = New-ScheduledTaskAction `
    -Execute $GitBash `
    -Argument "--login -c `"bash '$Script' >> '$LogFile' 2>&1`""

# トリガー: 毎日 18:00
$Trigger = New-ScheduledTaskTrigger -Daily -At "18:00"

# 設定
$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable

# タスク登録
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Claude Codeで作成したプログラムを毎日GitHubに自動保存" `
    -RunLevel Limited

Write-Host "タスクスケジューラへの登録が完了しました: $TaskName" -ForegroundColor Green
Write-Host "毎日 18:00 に自動でGitHubへコミット・プッシュします" -ForegroundColor Cyan
