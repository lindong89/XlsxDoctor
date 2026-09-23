<#
    XlsxDoctor 格式门禁测试

    验证「什么能进来、什么被挡住、挡住时说的是什么话」。
    这部分逻辑集中在 FormatSniffer.RejectReason()，GUI 和 CLI 共用同一份 ——
    所以这里同时测两条路，防止将来只改了一边。

    不需要 Excel，不需要网络。

    用法：
        .\tests\FormatGateTests.ps1
#>
[CmdletBinding()]
param(
    [string]$Exe
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Definition)
if (-not $Exe) {
    $Exe = Join-Path $root 'dist\XlsxDoctor-cli.exe'
    if (-not (Test-Path -LiteralPath $Exe)) { $Exe = Join-Path $root 'dist\XlsxDoctor.exe' }
}
if (-not (Test-Path -LiteralPath $Exe)) { throw "找不到可执行文件，请先运行 .\build.ps1" }
$Exe = (Resolve-Path -LiteralPath $Exe).Path
$guiExe = Join-Path $root 'dist\XlsxDoctor.exe'

$script:pass = 0
$script:fail = 0
$script:failedNames = @()

function Section([string]$t) {
    Write-Host ''
    Write-Host ("=== " + $t + " ===") -ForegroundColor Cyan
}
function Check([string]$n, [bool]$ok, [string]$d) {
    if ($ok) { $script:pass++; Write-Host ("  [PASS] " + $n) -ForegroundColor Green }
    else {
        $script:fail++; $script:failedNames += $n
        Write-Host ("  [FAIL] " + $n) -ForegroundColor Red
        if ($d) { Write-Host ("         " + $d) -ForegroundColor DarkRed }
    }
}

function Run([string[]]$a) {
    # 被拒绝的文件会往 stderr 写原因，而 $ErrorActionPreference='Stop' 会把原生程序的
    # stderr 当成异常抛出来。这里临时降级，否则"预期失败"的用例根本跑不到断言。
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & $Exe @a 2>&1
        return [pscustomobject]@{ Code = $LASTEXITCODE; Out = ($out | Out-String) }
    } finally { $ErrorActionPreference = $prev }
}

$work = Join-Path $env:TEMP ('xlsxdoctor-gate-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work | Out-Null

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---- 造一个最小但合法、且含垃圾的 xlsx（用来证明门禁没有误伤正常文件）----
function New-MinimalXlsx([string]$path) {
    $entries = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="KeepMe">Sheet1!$A$1</definedName><definedName name="BrokenName">#REF!</definedName></definedNames></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>10</v></c><c r="B1"><f>KeepMe*2</f><v>20</v></c></row></sheetData></worksheet>'
    }
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($k in $entries.Keys) {
            $e = $zip.CreateEntry($k, [System.IO.Compression.CompressionLevel]::Optimal)
            $s = $e.Open()
            try {
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($entries[$k])
                $s.Write($bytes, 0, $bytes.Length)
            } finally { $s.Dispose() }
        }
    } finally { $zip.Dispose() }
}

# ---- 造一个只有 CFBF 文件头的假 .xls（门禁只看文件头，不需要真能解析）----
function New-FakeXls([string]$path) {
    $bytes = New-Object byte[] 2048
    $head = [byte[]]@(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    [Array]::Copy($head, $bytes, 8)
    [System.IO.File]::WriteAllBytes($path, $bytes)
}

Write-Host ''
Write-Host 'XlsxDoctor 格式门禁测试' -ForegroundColor White
Write-Host ("exe : {0}" -f $Exe)
Write-Host ''

$realXlsx  = Join-Path $work 'real.xlsx'
New-MinimalXlsx $realXlsx

$fakeXls   = Join-Path $work 'real.xls'        # 内容是真 xlsx，名字叫 .xls
Copy-Item -LiteralPath $realXlsx -Destination $fakeXls -Force

$cfbfXls   = Join-Path $work 'cfbf.xls'        # 内容是真 CFBF，名字叫 .xls
New-FakeXls $cfbfXls

$cfbfNamedXlsx = Join-Path $work 'cfbf.xlsx'   # 内容是真 CFBF，名字却叫 .xlsx
New-FakeXls $cfbfNamedXlsx

$notExcel  = Join-Path $work 'notes.xlsx'
[System.IO.File]::WriteAllText($notExcel, 'this is just a text file, not a workbook', (New-Object System.Text.UTF8Encoding($false)))

$xlsb      = Join-Path $work 'book.xlsb'
[System.IO.File]::WriteAllText($xlsb, 'pretend xlsb', (New-Object System.Text.UTF8Encoding($false)))

# ---------------------------------------------------------- 1. .xls 被挡住
Section '1. .xls（CFBF/BIFF8）应当被挡住，并说清原因'
$r = Run @($cfbfXls)
Check '.xls 被拒绝（退出码非 0）' ($r.Code -ne 0) ("退出码 = " + $r.Code)
Check '提示里出现「暂不支持」' ($r.Out -match '暂不支持') $r.Out
Check '提示里点明是 .xls / BIFF8' ($r.Out -match '\.xls' -and $r.Out -match 'BIFF8') $r.Out
Check '提示里给出可操作的建议（另存为 .xlsx）' ($r.Out -match '另存为' -and $r.Out -match '\.xlsx') $r.Out
Check '被拒绝时没有产生输出文件' (-not (Test-Path -LiteralPath (Join-Path $work 'cfbf - 已清理.xls'))) '不该有输出'
Check '--dry-run 也照样拒绝' ((Run @($cfbfXls, '--dry-run')).Code -ne 0) '--dry-run 下没拒绝'

# ---------------------------------------------------------- 2. 按内容判定，不看扩展名
Section '2. 格式按内容判定，不看扩展名'
$r = Run @($fakeXls, '--dry-run')
Check '内容是真 xlsx 时，即使叫 .xls 也照常处理' ($r.Code -eq 0 -and $r.Out -match '分析报告') ("退出码 = " + $r.Code)
Check '并给出扩展名与内容不符的提示' ($r.Out -match '扩展名是 \.xls，但文件内容其实是 OOXML') $r.Out
Check '该提示只出现一次（不重复）' ((($r.Out | Select-String -Pattern '扩展名是 \.xls' -AllMatches).Matches.Count) -eq 1) `
      ("出现次数 = " + (($r.Out | Select-String -Pattern '扩展名是 \.xls' -AllMatches).Matches.Count))

$r = Run @($cfbfNamedXlsx)
Check '内容是真 CFBF 时，即使叫 .xlsx 也被挡住' ($r.Code -ne 0 -and $r.Out -match '暂不支持') ("退出码 = " + $r.Code)

# ---------------------------------------------------------- 3. 正常文件不受影响
Section '3. 正常 .xlsx 不受门禁影响'
$outOk = Join-Path $work 'real-out.xlsx'
$r = Run @($realXlsx, '-o', $outOk)
Check '真 xlsx 清理成功' ($r.Code -eq 0) ("退出码 = " + $r.Code + "`n" + $r.Out)
Check '输出文件已写出' (Test-Path -LiteralPath $outOk) '没有输出文件'
if (Test-Path -LiteralPath $outOk) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($outOk)
    try {
        $wb = $zip.Entries | Where-Object { $_.FullName -eq 'xl/workbook.xml' }
        $sr = New-Object System.IO.StreamReader($wb.Open())
        $text = $sr.ReadToEnd(); $sr.Dispose()
        Check '引用中的名称 KeepMe 保留' ($text -match 'name="KeepMe"') $text
        Check '失效名称 BrokenName 删除' (-not ($text -match 'BrokenName')) $text
    } finally { $zip.Dispose() }
}

# ---------------------------------------------------------- 4. 非 Excel 文件
Section '4. 非 Excel 文件'
$r = Run @($notExcel)
Check '纯文本被拒绝' ($r.Code -ne 0) ("退出码 = " + $r.Code)
Check '提示里说明无法识别格式' ($r.Out -match '无法识别的文件格式') $r.Out
Check '提示里指出它可能不是 Excel 文件' ($r.Out -match '不是 Excel 文件|已经损坏') $r.Out

$r = Run @($xlsb)
Check '.xlsb 被拒绝' ($r.Code -ne 0) ("退出码 = " + $r.Code)
Check '.xlsb 的提示不误导（不说是 .xls）' ($r.Out -notmatch 'BIFF8') $r.Out

# ---------------------------------------------------------- 5. GUI 也要挡住
Section '5. GUI 走同一套判定'
if (-not (Test-Path -LiteralPath $guiExe)) {
    Check 'GUI 端到端（找不到 exe，跳过）' $true ''
} else {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class GateUiBot
{
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    public static List<string> Visible()
    {
        List<string> rows = new List<string>();
        EnumWindows(delegate(IntPtr h, IntPtr p)
        {
            if (!IsWindowVisible(h)) return true;
            StringBuilder t = new StringBuilder(512); GetWindowTextW(h, t, 512);
            StringBuilder c = new StringBuilder(256); GetClassNameW(h, c, 256);
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (t.Length > 0) rows.Add(pid + "|" + c.ToString() + "|" + t.ToString());
            return true;
        }, IntPtr.Zero);
        return rows;
    }
}
'@
    function Get-GuiWindows([int]$procId) {
        $rows = @()
        foreach ($row in [GateUiBot]::Visible()) {
            $parts = $row -split '\|', 3
            if ([int]$parts[0] -eq $procId) { $rows += [pscustomobject]@{ Class = $parts[1]; Title = $parts[2] } }
        }
        return $rows
    }

    # 轮询等待，不要用固定 Sleep —— 机器忙的时候 6 秒可能刚好不够，
    # 测试就会随机红一次，那种红是噪声不是信号。
    function Wait-GuiWindow([int]$procId, [string]$classLike, [int]$timeoutSec = 20) {
        $deadline = (Get-Date).AddSeconds($timeoutSec)
        while ((Get-Date) -lt $deadline) {
            $w = @(Get-GuiWindows $procId | Where-Object { $_.Class -like $classLike })
            if ($w.Count -gt 0) { return $w }
            Start-Sleep -Milliseconds 400
        }
        return @()
    }

    Get-Process -Name 'XlsxDoctor*' -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 800

    $proc = Start-Process -FilePath $guiExe -ArgumentList ('"' + $cfbfXls + '"') -PassThru
    # 等对话框出现；$wins 始终是「全部窗口」，$dlg 单独放对话框，
    # 否则后面「主窗口仍在」的断言会因为 $wins 里只有对话框而拿不到主窗口
    $dlgWins = @(Wait-GuiWindow $proc.Id '#32770')
    $wins = @(Get-GuiWindows $proc.Id)
    $dlg = $dlgWins | Select-Object -First 1
    Check 'GUI 弹出了对话框' ($null -ne $dlg) (($wins | ForEach-Object { $_.Class + '/' + $_.Title }) -join ' ; ')
    Check '对话框标题是「暂不支持」' ($null -ne $dlg -and $dlg.Title -eq '暂不支持') $(if ($dlg) { $dlg.Title } else { '没有对话框' })
    $main = $wins | Where-Object { $_.Class -like 'WindowsForms10*' } | Select-Object -First 1
    Check '主窗口仍然在（没有崩溃退出）' ($null -ne $main) '主窗口不见了'

    # 关掉对话框和主窗口
    Get-Process -Name 'XlsxDoctor*' -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 500

    # 对照：真 xlsx 不应该弹对话框
    $proc2 = Start-Process -FilePath $guiExe -ArgumentList ('"' + $realXlsx + '"') -PassThru
    $main2 = @(Wait-GuiWindow $proc2.Id 'WindowsForms10*')
    Check '真 xlsx 打开了主窗口' ($main2.Count -gt 0) '没有主窗口（轮询 20 秒）'
    # 主窗口出现后再多等一会，给可能的对话框留出时间（它可能比主窗口晚弹）
    Start-Sleep -Seconds 3
    $wins2 = Get-GuiWindows $proc2.Id
    $dlg2 = $wins2 | Where-Object { $_.Class -eq '#32770' } | Select-Object -First 1
    Check '真 xlsx 打开时不弹任何对话框' ($null -eq $dlg2) $(if ($dlg2) { $dlg2.Title } else { '' })
    Get-Process -Name 'XlsxDoctor*' -ErrorAction SilentlyContinue | Stop-Process -Force
}

Write-Host ''
Write-Host '========================================' -ForegroundColor White
Write-Host (" 通过 {0} / 失败 {1}" -f $script:pass, $script:fail) -ForegroundColor White
if ($script:fail -gt 0) {
    Write-Host ' 失败项：' -ForegroundColor Red
    foreach ($n in $script:failedNames) { Write-Host ("   - " + $n) -ForegroundColor Red }
}
Write-Host (" 工作目录: " + $work) -ForegroundColor DarkGray
Write-Host '========================================' -ForegroundColor White

if ($script:fail -gt 0) { exit 1 } else { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue; exit 0 }
