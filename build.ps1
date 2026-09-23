<#
    XlsxDoctor 一键编译脚本
    只依赖 .NET Framework 4.x 自带的 csc.exe —— 不需要装 Visual Studio / .NET SDK。

    会产出两个 exe（同一份源码，只是子系统不同）：
        dist\XlsxDoctor.exe       图形界面，双击运行，不闪黑框
        dist\XlsxDoctor-cli.exe   控制台程序，给批处理 / 脚本用

    用法：
        .\build.ps1                  编译到 .\dist\
        .\build.ps1 -DebugBuild      带调试符号编译
        .\build.ps1 -OutDir C:\x     指定输出目录
#>
[CmdletBinding()]
param(
    [string]$OutDir,
    [switch]$DebugBuild
)

$ErrorActionPreference = 'Stop'

$root   = Split-Path -Parent $MyInvocation.MyCommand.Definition
$srcDir = Join-Path $root 'src'
$manifest = Join-Path $root 'app.manifest'

if (-not $OutDir) { $OutDir = Join-Path $root 'dist' }
if (-not (Test-Path -LiteralPath $srcDir)) { throw "找不到源码目录: $srcDir" }

Write-Host ''
Write-Host '=== XlsxDoctor 编译 ===' -ForegroundColor Cyan
Write-Host ''

# ---------------------------------------------------------------- 1. 找 csc.exe
$cscCandidates = @(
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'),
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe')
)
$csc = $null
foreach ($c in $cscCandidates) { if (Test-Path -LiteralPath $c) { $csc = $c; break } }
if (-not $csc) {
    throw '找不到 csc.exe。本机需要 .NET Framework 4.x（Windows 10/11 默认自带）。'
}
Write-Host ("编译器 : {0}" -f $csc)
Write-Host ("          {0}" -f (Get-Item -LiteralPath $csc).VersionInfo.FileVersion)

# ------------------------------------------------- 2. 确保 .cs 是 UTF-8 带 BOM
# csc.exe 读源码时若没有 BOM，会按系统 ANSI 代码页解码 —— 中文会直接乱掉甚至编译失败。
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$utf8Raw = New-Object System.Text.UTF8Encoding($false)
$bomFixed = 0
Get-ChildItem -LiteralPath $srcDir -Filter *.cs -Recurse | ForEach-Object {
    $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    if (-not $hasBom) {
        $text = $utf8Raw.GetString($bytes)
        [System.IO.File]::WriteAllText($_.FullName, $text, $utf8Bom)
        $bomFixed++
        Write-Host ("  [BOM] {0}" -f $_.Name) -ForegroundColor DarkYellow
    }
}
if ($bomFixed -eq 0) { Write-Host '编码   : 全部 .cs 已是 UTF-8 带 BOM' }
else                 { Write-Host ("编码   : 已为 {0} 个文件补上 BOM" -f $bomFixed) }

# -------------------------------------------------------------------- 3. 图标
# csc.exe 的 /win32icon 只认 .ico，不认 .png。assets 里没有就现场从 PNG 生成一个。
$iconFile = Join-Path $root 'assets\XlsxDoctor.ico'
if (-not (Test-Path -LiteralPath $iconFile)) {
    $mk = Join-Path $root 'tools\make-icon.ps1'
    if (Test-Path -LiteralPath $mk) {
        Write-Host '图标   : assets\XlsxDoctor.ico 不存在，正在从 PNG 生成…'
        & $mk | Out-Null
    }
}
$hasIcon = Test-Path -LiteralPath $iconFile
if ($hasIcon) {
    Write-Host ('图标   : {0}（{1:N0} 字节）' -f (Split-Path $iconFile -Leaf), (Get-Item -LiteralPath $iconFile).Length)
} else {
    Write-Host '图标   : 未找到 assets\XlsxDoctor.ico，产物将使用系统默认图标' -ForegroundColor DarkYellow
}
Write-Host ''

# -------------------------------------------------------------------- 4. 编译
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$refs = @(
    'System.dll'
    'System.Core.dll'
    'System.Windows.Forms.dll'
    'System.Drawing.dll'
    'System.IO.Compression.dll'
    'System.IO.Compression.FileSystem.dll'
    'System.Xml.dll'
)

$sources = @(Get-ChildItem -LiteralPath $srcDir -Filter *.cs -Recurse | Sort-Object FullName)
Write-Host ("源文件 : {0} 个" -f $sources.Count)
Write-Host ''

$targets = @(
    [pscustomobject]@{ File = 'XlsxDoctor.exe';     Target = 'winexe'; Note = '图形界面（双击运行）' }
    [pscustomobject]@{ File = 'XlsxDoctor-cli.exe'; Target = 'exe';    Note = '命令行 / 批处理' }
)

$results = @()
$failed = 0

foreach ($t in $targets) {
    $outExe = Join-Path $OutDir $t.File

    $cscArgs = New-Object System.Collections.Generic.List[string]
    $cscArgs.Add('/nologo')
    $cscArgs.Add(('/target:{0}' -f $t.Target))
    $cscArgs.Add('/platform:anycpu')
    $cscArgs.Add('/utf8output')
    if ($DebugBuild) { $cscArgs.Add('/debug+'); $cscArgs.Add('/optimize-') }
    else             { $cscArgs.Add('/debug-'); $cscArgs.Add('/optimize+') }
    $cscArgs.Add(('/out:{0}' -f $outExe))
    if (Test-Path -LiteralPath $manifest) { $cscArgs.Add(('/win32manifest:{0}' -f $manifest)) }
    if ($hasIcon) { $cscArgs.Add(('/win32icon:{0}' -f $iconFile)) }
    foreach ($r in $refs) { $cscArgs.Add(('/r:{0}' -f $r)) }
    foreach ($s in $sources) { $cscArgs.Add($s.FullName) }

    Write-Host ("-> {0}  ({1})" -f $t.File, $t.Note) -ForegroundColor DarkGray

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $cscOut = & $csc $cscArgs.ToArray() 2>&1
    $code = $LASTEXITCODE
    $sw.Stop()

    if ($code -ne 0) {
        Write-Host ("   编译失败（csc 退出码 {0}）" -f $code) -ForegroundColor Red
        foreach ($line in $cscOut) { Write-Host ("     {0}" -f $line) -ForegroundColor DarkRed }

        # 最常见的失败原因：上一次运行的 GUI 还开着，锁住了 exe
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($t.File)
        $running = @(Get-Process -Name $stem -ErrorAction SilentlyContinue)
        if ($running.Count -gt 0) {
            Write-Host ("     提示：{0} 正在运行（PID {1}）。先关掉它再编译，或执行：" -f $t.File, (($running | ForEach-Object { $_.Id }) -join ', ')) -ForegroundColor Yellow
            Write-Host ("           Get-Process -Name '{0}' | Stop-Process -Force" -f $stem) -ForegroundColor Yellow
        }
        $failed++
    }
    elseif (-not (Test-Path -LiteralPath $outExe)) {
        Write-Host '   编译未报错但没有产出 exe。' -ForegroundColor Red
        $failed++
    }
    else {
        $exe = Get-Item -LiteralPath $outExe
        Write-Host ("   成功  {0:N0} 字节  {1:N1} 秒" -f $exe.Length, $sw.Elapsed.TotalSeconds) -ForegroundColor Green
        $results += $exe
    }
}

Write-Host ''
if ($failed -gt 0) {
    Write-Host ("编译失败：{0} 个目标未产出。" -f $failed) -ForegroundColor Red
    exit 1
}

Write-Host '全部编译成功。' -ForegroundColor Green
Write-Host ''
foreach ($e in $results) {
    Write-Host ("  {0,-20} {1,10:N0} 字节   版本 {2}" -f $e.Name, $e.Length, $e.VersionInfo.FileVersion)
}
Write-Host ''
Write-Host '用法：'
Write-Host '  双击 dist\XlsxDoctor.exe            -> 图形界面（也可把 xlsx 拖到 exe 上）'
Write-Host '  dist\XlsxDoctor-cli.exe --help      -> 命令行帮助'
Write-Host '  dist\XlsxDoctor-cli.exe 报表.xlsx   -> 清理并输出 "报表 - 已清理.xlsx"'
Write-Host ''
