param(
    [string]$RemoteUrl = "https://github.com/duybao2812/Add-in-Word.git",
    [string]$CommitMessage = "feat: WinForms Classic Word Add-in with OCR engine and hardware detection profile"
)

$gitExe = "C:\Users\Admin\AppData\Local\GitHubDesktop\app-3.6.5\resources\app\git\cmd\git.exe"

if (-not (Test-Path $gitExe)) {
    Write-Error "Khong tim thay Git tai: $gitExe"
    exit 1
}

Write-Host ">>> Khoi tao / kiem tra Git repository..."
if (-not (Test-Path ".git")) {
    & $gitExe init
}

& $gitExe config user.name "duybao2812"
& $gitExe config user.email "gunfire2812@gmail.com"
& $gitExe branch -M main

Write-Host ">>> Staging cac tep tin..."
& $gitExe add .

Write-Host ">>> Danh sach tep staged:"
& $gitExe status --short

Write-Host ">>> Tao commit..."
& $gitExe commit -m $CommitMessage

Write-Host ">>> Thiet lap remote origin: $RemoteUrl"
$remotes = & $gitExe remote
if ($remotes -contains "origin") {
    & $gitExe remote set-url origin $RemoteUrl
} else {
    & $gitExe remote add origin $RemoteUrl
}

Write-Host ">>> Tien hanh day code len GitHub (git push -u origin main)..."
& $gitExe push -u origin main
