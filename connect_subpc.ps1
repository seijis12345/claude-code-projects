$cred = Get-Credential -UserName 'keysight\seijis' -Message 'サブPC (5CD323LLG0) のパスワードを入力してください'
$pw = $cred.GetNetworkCredential().Password
$result = net use \\5CD323LLG0\C$ $pw /USER:seijis 2>&1
Write-Output $result

if ($LASTEXITCODE -eq 0) {
    Write-Output "接続成功"
    # .py files from .spyder-py3
    $src = "\\5CD323LLG0\C$\Users\seijis\.spyder-py3"
    $dst = "C:\Users\seijis\OneDrive - Keysight Technologies\PythonCode"

    $files = Get-ChildItem -Path $src -Filter "*.py" | Where-Object { $_.Name -notin @("temp.py","history.py","history_internal.py") }
    foreach ($f in $files) {
        $dstFile = Join-Path $dst $f.Name
        if (Test-Path $dstFile) {
            $newName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name) + "_sub" + $f.Extension
            $dstFile = Join-Path $dst $newName
            Write-Output "重複のため名前変更: $($f.Name) -> $newName"
        } else {
            Write-Output "コピー: $($f.Name)"
        }
        Copy-Item $f.FullName $dstFile
    }

    # .py files from home dir
    $homeFiles = Get-ChildItem -Path "\\5CD323LLG0\C$\Users\seijis" -Filter "*.py" -ErrorAction SilentlyContinue
    foreach ($f in $homeFiles) {
        $dstFile = Join-Path $dst $f.Name
        if (Test-Path $dstFile) {
            $newName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name) + "_sub" + $f.Extension
            $dstFile = Join-Path $dst $newName
            Write-Output "重複のため名前変更: $($f.Name) -> $newName"
        } else {
            Write-Output "コピー: $($f.Name)"
        }
        Copy-Item $f.FullName $dstFile
    }

    net use \\5CD323LLG0\C$ /delete | Out-Null
    Write-Output "完了"
} else {
    Write-Output "接続失敗: $result"
}
