# PSI agent 一键安装（Windows）。由 “一键安装_Windows.bat” 调用，不用直接运行。
# 做的事：找到或安装 Python 3.11+ → 建 .venv → 装依赖 → 建 数据 文件夹 → 检查能不能用。
# 重复运行没关系，已经装好的步骤会跳过。

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$env:PYTHONUTF8 = '1'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

function Say($msg, $color = 'White') { Write-Host $msg -ForegroundColor $color }
function Fail($msg) { Say "`n[没装好] $msg" 'Red'; Say '把这个窗口截图发给负责人。' 'Red'; exit 1 }

Say "== PSI agent 一键安装 ==" 'Cyan'
Say "安装位置：$Root`n"

# 1. 找一个 3.11 以上的 Python
function Find-Python {
    $check = 'import sys; assert sys.version_info >= (3, 11); print(sys.executable)'
    foreach ($v in '3.13', '3.12', '3.11') {
        try { $p = & py "-$v" -c $check 2>$null; if ($LASTEXITCODE -eq 0 -and $p) { return "$p".Trim() } } catch {}
    }
    foreach ($c in 'python', 'python3') {
        try { $p = & $c -c $check 2>$null; if ($LASTEXITCODE -eq 0 -and $p) { return "$p".Trim() } } catch {}
    }
    foreach ($v in '313', '312', '311') {
        $p = Join-Path $env:LOCALAPPDATA "Programs\Python\Python$v\python.exe"
        if (Test-Path $p) { return $p }
    }
    return $null
}

Say "[1/4] 检查 Python ..."
$Py = $null
if (Test-Path '.venv\Scripts\python.exe') {
    & '.venv\Scripts\python.exe' -c 'import sys; assert sys.version_info >= (3, 11)' 2>$null
    if ($LASTEXITCODE -ne 0) { Say '      旧的 .venv 不能用，重建。' 'Yellow'; Remove-Item -Recurse -Force '.venv' }
}
if (-not (Test-Path '.venv\Scripts\python.exe')) {
    $Py = Find-Python
    if (-not $Py) {
        Say '      没找到 Python 3.11 以上，尝试自动安装 Python 3.13（只装给当前用户，不用管理员）...' 'Yellow'
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Say '      如果下面问是否同意条款，输入 Y 回车。' 'Yellow'
            winget install -e --id Python.Python.3.13 --scope user
            $Py = Find-Python
        }
        if (-not $Py) {
            Start-Process 'https://www.python.org/downloads/windows/'
            Fail ("自动安装没成功（公司电脑可能限制了 winget）。`n" +
                  "已打开 Python 下载页：下载 Python 3.13 的 Windows installer，安装时一定勾选 'Add python.exe to PATH'，" +
                  "装完再双击一次 一键安装_Windows.bat。装不了的话找 IT。")
        }
    }
    Say "      用这个 Python：$Py" 'Green'
} else {
    Say '      .venv 已经有了，跳过。' 'Green'
}

# 2. 建虚拟环境
Say "[2/4] 建虚拟环境 .venv ..."
if (-not (Test-Path '.venv\Scripts\python.exe')) {
    & $Py -m venv .venv
    if (-not (Test-Path '.venv\Scripts\python.exe')) { Fail '建 .venv 失败。' }
}
$VPy = Join-Path $Root '.venv\Scripts\python.exe'
Say '      好了。' 'Green'

# 3. 装依赖（官方源不通就换镜像）
Say "[3/4] 安装依赖（第一次大约 1–3 分钟）..."
& $VPy -m pip install --upgrade pip -q 2>$null
& $VPy -m pip install -r requirements.txt -q
if ($LASTEXITCODE -ne 0) {
    Say '      官方源没装上，换清华镜像再试 ...' 'Yellow'
    & $VPy -m pip install -r requirements.txt -q -i https://pypi.tuna.tsinghua.edu.cn/simple
    if ($LASTEXITCODE -ne 0) { Fail '依赖没装上，多半是网络或公司代理拦了。' }
}
& $VPy -c 'import pandas, openpyxl, yaml, xlrd'
if ($LASTEXITCODE -ne 0) { Fail '依赖装完但导入失败。' }
Say '      好了。' 'Green'

# 4. 数据文件夹
Say "[4/4] 准备 数据 文件夹 ..."
New-Item -ItemType Directory -Force -Path '数据' | Out-Null
$Psi = Join-Path $Root '数据\Poland PSI DB+Channel.xlsx'

Say "`n== 安装完成 ==" 'Green'
if (Test-Path $Psi) {
    Say '数据\Poland PSI DB+Channel.xlsx 已经在了，可以直接双击 生成看板_Windows.bat。' 'Green'
} else {
    Say '还差一步：从飞书导出 PSI 表，改名为  Poland PSI DB+Channel.xlsx  放进刚打开的 数据 文件夹。' 'Yellow'
    Say '然后双击 生成看板_Windows.bat 就能出看板。' 'Yellow'
    Start-Process explorer.exe (Join-Path $Root '数据')
}
exit 0
