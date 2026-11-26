# 生成 PSI 看板（Windows）。由 “生成看板_Windows.bat” 调用。
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$env:PYTHONUTF8 = '1'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

function Say($msg, $color = 'White') { Write-Host $msg -ForegroundColor $color }

$VPy = Join-Path $Root '.venv\Scripts\python.exe'
if (-not (Test-Path $VPy)) {
    Say '还没安装，先自动跑一遍安装 ...' 'Yellow'
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'setup_windows.ps1')
    if (-not (Test-Path $VPy)) { exit 1 }
}
$Psi = Join-Path $Root '数据\Poland PSI DB+Channel.xlsx'
if (-not (Test-Path $Psi)) {
    Say '找不到 数据\Poland PSI DB+Channel.xlsx。' 'Red'
    Say '从飞书导出 PSI 表，改名为  Poland PSI DB+Channel.xlsx  放进 数据 文件夹后再双击。' 'Red'
    New-Item -ItemType Directory -Force -Path '数据' | Out-Null
    Start-Process explorer.exe (Join-Path $Root '数据')
    exit 1
}

$week = Read-Host '生成第几周的看板？输入数字，比如 47（直接回车 = 最近填好的一周）'
$args2 = @('-m', 'tools.dashboard')
if ($week -match '^\d+$') { $args2 += @('--week', $week) }
Say '正在生成（大约 1 分钟）...'
$out = & $VPy @args2 2>$null
if ($LASTEXITCODE -ne 0 -or -not $out) {
    Say '生成失败，下面是详细信息：' 'Red'
    & $VPy @args2
    exit 1
}
$html = ($out | Select-Object -Last 1).ToString().Trim()
Say "生成好了：$html" 'Green'
Say '看板里是 PSI 实数（L3 机密），只在公司渠道传。' 'Yellow'
if (Test-Path $html) { Start-Process $html }
exit 0
