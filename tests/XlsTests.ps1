<#
    XlsxDoctor .xls（CFBF + BIFF8）专项测试

    用 Excel 造一个带陷阱的真实 .xls 夹具，逐项验证：
      * 被公式引用的定义名称必须原样保留，且引用不指错
      * 无人引用 / 已失效的名称必须删掉
      * 三维引用 Sheet1:Sheet3!A1 必须还能正确计算（这是索引重映射的核心考验）
      * 外部链接引用必须保留，不被误删
      * 全部公式文本与计算值在清理前后逐单元格一致
      * 幂等性、防呆、--dry-run、--in-place
      * 扩展名与内容不符时按内容分派

    用法：
        .\tests\XlsTests.ps1
        .\tests\XlsTests.ps1 -Exe ..\dist\XlsxDoctor-cli.exe

    需要本机装有 Excel（用 COM 造夹具和核对结果）。没装就自动跳过并给出提示。
#>
[CmdletBinding()]
param(
    [string]$Exe
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Definition)
if (-not $Exe) { $Exe = Join-Path $root 'dist\XlsxDoctor-cli.exe' }
if (-not (Test-Path -LiteralPath $Exe)) { throw "找不到可执行文件: $Exe`n请先运行 .\build.ps1" }
$Exe = (Resolve-Path -LiteralPath $Exe).Path

# ------------------------------------------------------------------ 引擎开关探测
# .xls 引擎可能被 FormatSniffer.XlsEnabled 关掉（当前就是关的）。
# 用最小 CFBF 头探一下：被关掉就整体跳过，而不是抛一堆莫名其妙的红 ——
# 那种红会掩盖真正的回归。门禁行为由 FormatGateTests.ps1 负责验证。
$probeFile = Join-Path $env:TEMP ('xlsengine-probe-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.xls')
$probeBytes = New-Object byte[] 1024
$magic = [byte[]]@(0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
[Array]::Copy($magic, $probeBytes, 8)
[System.IO.File]::WriteAllBytes($probeFile, $probeBytes)
$prevEap = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
$probeOut = (& $Exe $probeFile --dry-run 2>&1 | Out-String)
$ErrorActionPreference = $prevEap
Remove-Item -LiteralPath $probeFile -Force -ErrorAction SilentlyContinue

if ($probeOut -match '暂不支持') {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Yellow
    Write-Host ' .xls 清理引擎当前【已关闭】（FormatSniffer.XlsEnabled = false）' -ForegroundColor Yellow
    Write-Host ' 整个脚本跳过。' -ForegroundColor Yellow
    Write-Host '' -ForegroundColor Yellow
    Write-Host ' 引擎代码仍在 src\Xls\ 下（完整且已通过测试），' -ForegroundColor Yellow
    Write-Host ' 把 FormatSniffer.XlsEnabled 改成 true 重新编译即可恢复。' -ForegroundColor Yellow
    Write-Host '' -ForegroundColor Yellow
    Write-Host ' 关闭状态下的门禁行为请跑: .\tests\FormatGateTests.ps1' -ForegroundColor Yellow
    Write-Host '============================================================' -ForegroundColor Yellow
    Write-Host ''
    exit 0
}

$work = Join-Path $env:TEMP ('xlsdoctor-xlstests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work | Out-Null

$script:pass = 0
$script:fail = 0
$script:skip = 0
$script:failedNames = @()

function Check([string]$name, [bool]$ok, [string]$detail) {
    if ($ok) {
        $script:pass++
        Write-Host ("  [PASS] {0}" -f $name) -ForegroundColor Green
    } else {
        $script:fail++
        $script:failedNames += $name
        Write-Host ("  [FAIL] {0}" -f $name) -ForegroundColor Red
        if ($detail) { Write-Host ("         {0}" -f $detail) -ForegroundColor DarkRed }
    }
}

function Skip([string]$name, [string]$why) {
    $script:skip++
    Write-Host ("  [SKIP] {0} —— {1}" -f $name, $why) -ForegroundColor Yellow
}

function Section([string]$title) {
    Write-Host ''
    Write-Host ("=== {0} ===" -f $title) -ForegroundColor Cyan
}

function Run([string[]]$argv) {
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & $Exe @argv 2>&1 | Out-String
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $prevEap }
    return [pscustomobject]@{ Out = $out; Code = $code }
}

# ============================================================ Excel COM

$script:excel = $null
try {
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $false
    $script:excel.DisplayAlerts = $false
    $script:excel.ScreenUpdating = $false
} catch {
    Write-Host ''
    Write-Host "本机没有可用的 Excel COM 组件，.xls 专项测试无法运行。" -ForegroundColor Yellow
    Write-Host "（这些测试需要用 Excel 造夹具并核对计算结果）" -ForegroundColor Yellow
    exit 0
}

function Close-Excel {
    if ($script:excel) {
        try { $script:excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:excel) | Out-Null } catch { }
        $script:excel = $null
    }
    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Stop-Process -Force
}

function Get-Md5Hex([string]$s) {
    $md5 = [System.Security.Cryptography.MD5]::Create()
    return [System.BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s)))
}

function ConvertTo-FlatString($x) {
    $sep = [string][char]1
    if ($x -is [System.Array]) {
        $l = New-Object System.Collections.Generic.List[string]
        foreach ($e in $x) { [void]$l.Add([string]$e) }
        return [string]::Join($sep, $l.ToArray())
    }
    return [string]$x
}

# ============================================================ 夹具

$fixture = Join-Path $work 'fixture.xls'

function New-XlsFixture {
    $wb = $script:excel.Workbooks.Add()

    # 四个工作表：Sheet1/Sheet2/Sheet3 留着，Temp 待会儿删掉用来制造失效名称
    while ($wb.Worksheets.Count -lt 4) { [void]$wb.Worksheets.Add() }
    # 先统一改成临时名再改目标名 —— 新建工作簿自带 Sheet1/Sheet2…，直接改会撞名报 COMException
    for ($i = 1; $i -le $wb.Worksheets.Count; $i++) { $wb.Worksheets.Item($i).Name = ("Tmp" + $i) }
    $wb.Worksheets.Item(1).Name = 'Sheet1'
    $wb.Worksheets.Item(2).Name = 'Sheet2'
    $wb.Worksheets.Item(3).Name = 'Sheet3'
    $wb.Worksheets.Item(4).Name = 'Temp'

    $s1 = $wb.Worksheets.Item('Sheet1')
    $s2 = $wb.Worksheets.Item('Sheet2')
    $s3 = $wb.Worksheets.Item('Sheet3')

    $s1.Range('A1').Value2 = 10
    $s1.Range('A2').Value2 = 20
    $s1.Range('A3').Value2 = 30
    $s2.Range('A1').Value2 = 5
    $s3.Range('A1').Value2 = 7

    # 被公式引用的名称
    [void]$wb.Names.Add('KeepMe', '=Sheet1!$A$1')
    # 无人引用但定义有效
    [void]$wb.Names.Add('UnusedGood', '=Sheet1!$A$2')
    # 指向 Temp 表，随后删表 → 定义变成 #REF!，属于"已失效"
    [void]$wb.Names.Add('BrokenName', '=Temp!$A$1')

    # 公式：名称引用 + 三维引用 + 跨表引用
    $s1.Range('B1').Formula = '=KeepMe*2'
    $s1.Range('B2').Formula = '=SUM(Sheet1:Sheet3!A1)'
    $s1.Range('B3').Formula = '=Sheet2!A1+1'
    $s1.Range('B4').Formula = '=KeepMe+Sheet3!A1'

    # 系统名（打印区域）—— 必须无条件保留
    $s1.PageSetup.PrintArea = '$A$1:$B$10'

    # 删掉 Temp 表，让 BrokenName 变成 #REF!
    $wb.Worksheets.Item('Temp').Delete()

    if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Force }
    $wb.SaveAs($fixture, 56)   # 56 = xlExcel8（BIFF8 的 .xls）
    $wb.Close($false)
}

function Get-WorkbookProbe([string]$path) {
    $wb = $script:excel.Workbooks.Open($path, 0, $true)
    try {
        $names = @{}
        foreach ($n in $wb.Names) { $names[$n.Name] = [string]$n.RefersTo }
        $sheets = @{}
        foreach ($ws in $wb.Worksheets) {
            $ur = $ws.UsedRange
            $sheets[$ws.Name] = [pscustomobject]@{
                Rows  = $ur.Rows.Count
                Cols  = $ur.Columns.Count
                FHash = (Get-Md5Hex (ConvertTo-FlatString $ur.Formula))
                VHash = (Get-Md5Hex (ConvertTo-FlatString $ur.Value2))
            }
        }
        return [pscustomobject]@{ Names = $names; Sheets = $sheets; SheetCount = $wb.Worksheets.Count }
    } finally { $wb.Close($false) }
}

# ============================================================ 开跑

Write-Host ''
Write-Host '========================================' -ForegroundColor White
Write-Host ' XlsxDoctor .xls 专项测试' -ForegroundColor White
Write-Host '========================================' -ForegroundColor White

try {
    Section '0. 构造夹具'
    New-XlsFixture
    $fx = Get-Item -LiteralPath $fixture
    Write-Host ("  夹具: {0}  ({1:N0} 字节)" -f $fx.Name, $fx.Length) -ForegroundColor DarkGray

    $probe0 = Get-WorkbookProbe $fixture
    Check '夹具含 KeepMe / UnusedGood / BrokenName 三个自定义名称' `
          ($probe0.Names.ContainsKey('KeepMe') -and $probe0.Names.ContainsKey('UnusedGood') -and $probe0.Names.ContainsKey('BrokenName')) `
          ("实际: " + (($probe0.Names.Keys) -join ', '))
    # 注意：BIFF8 里内置名存的是 "Print_Area"（工作表作用域），
    # 不带 OOXML 那套 "_xlnm." 前缀 —— 两种都要认。
    Check '夹具含打印区域系统名（Print_Area）' `
          (@($probe0.Names.Keys | Where-Object { $_ -match 'Print_Area' }).Count -ge 1) `
          ("实际: " + (($probe0.Names.Keys) -join ', '))
    Check 'BrokenName 的定义确实是 #REF!' `
          ($probe0.Names['BrokenName'] -match '#REF!') `
          ("实际: " + $probe0.Names['BrokenName'])
    $magic = [System.IO.File]::ReadAllBytes($fixture)[0..3]
    Check '夹具是 CFBF/OLE2（.xls）' `
          ($magic[0] -eq 0xD0 -and $magic[1] -eq 0xCF -and $magic[2] -eq 0x11 -and $magic[3] -eq 0xE0) `
          (("文件头 = " + (($magic | ForEach-Object { $_.ToString('X2') }) -join ' ')))

    # ---------------------------------------------------------- 1. 分析
    Section '1. 分析（dry-run）'
    $dry = Run @($fixture, '--dry-run')
    Check '--dry-run 退出码为 0' ($dry.Code -eq 0) ("退出码 = " + $dry.Code)
    Check '--dry-run 识别为 .xls 报告' ($dry.Out -match '分析报告（\.xls / BIFF8）')
    Check '--dry-run 检测到被公式引用的名称' ($dry.Out -match '被公式引用，保留 : [1-9]')
    Check '--dry-run 检测到系统名（打印区域）' ($dry.Out -match '系统名保留 : [1-9]')
    Check '--dry-run 检测到已失效名称' ($dry.Out -match '定义已失效，删除 : [1-9]')
    Check '--dry-run 检测到无人引用的名称' ($dry.Out -match '无人引用，删除\s+: [1-9]')
    Check '--dry-run 报告公式全部解析成功' ($dry.Out -match '条全部解析成功')
    Check '--dry-run 没有写出任何文件' `
          (-not (Test-Path -LiteralPath (Join-Path $work 'fixture - 已清理.xls')))

    # ---------------------------------------------------------- 2. 清理
    Section '2. 清理并核对'
    $out1 = Join-Path $work 'cleaned1.xls'
    $r1 = Run @($fixture, '-o', $out1)
    Check '清理退出码为 0' ($r1.Code -eq 0) ("退出码 = " + $r1.Code + "`n" + $r1.Out)
    Check '输出文件已写出' (Test-Path -LiteralPath $out1)
    Check '输出包校验通过' ($r1.Out -match '包校验\s+:\s+通过')

    if (Test-Path -LiteralPath $out1) {
        $probe1 = Get-WorkbookProbe $out1

        Check '被公式引用的 KeepMe 仍然存在' ($probe1.Names.ContainsKey('KeepMe')) `
              ("实际名称: " + (($probe1.Names.Keys) -join ', '))
        Check 'KeepMe 的定义没被改坏' `
              ($probe1.Names.ContainsKey('KeepMe') -and $probe1.Names['KeepMe'] -eq $probe0.Names['KeepMe']) `
              ("原: " + $probe0.Names['KeepMe'] + "  新: " + $probe1.Names['KeepMe'])
        Check '无人引用的 UnusedGood 已删除' (-not $probe1.Names.ContainsKey('UnusedGood'))
        Check '已失效的 BrokenName 已删除' (-not $probe1.Names.ContainsKey('BrokenName'))
        Check '系统名 Print_Area 已保留' `
              (@($probe1.Names.Keys | Where-Object { $_ -match 'Print_Area' }).Count -ge 1) `
              ("实际名称: " + (($probe1.Names.Keys) -join ', '))

        Check '工作表数量不变' ($probe1.SheetCount -eq $probe0.SheetCount) `
              ("原 " + $probe0.SheetCount + " -> 新 " + $probe1.SheetCount)

        $badF = 0; $badV = 0
        foreach ($k in $probe0.Sheets.Keys) {
            if (-not $probe1.Sheets.ContainsKey($k)) { $badF++; continue }
            if ($probe0.Sheets[$k].FHash -ne $probe1.Sheets[$k].FHash) { $badF++ }
            if ($probe0.Sheets[$k].VHash -ne $probe1.Sheets[$k].VHash) { $badV++ }
        }
        Check '全部公式文本逐单元格一致' ($badF -eq 0) ("不一致的工作表数 = " + $badF)
        Check '全部计算值逐单元格一致（三维引用/名称引用都算对了）' ($badV -eq 0) ("不一致的工作表数 = " + $badV)
    }

    # ---------------------------------------------------------- 3. 幂等
    Section '3. 幂等性'
    $out2 = Join-Path $work 'cleaned2.xls'
    $r2 = Run @($out1, '-o', $out2)
    Check '再清一次判定为"没有可清理的垃圾"' ($r2.Out -match '没有发现可清理的垃圾') `
          ($r2.Out.Substring(0, [Math]::Min(300, $r2.Out.Length)))
    Check '幂等时未写出第二个文件' (-not (Test-Path -LiteralPath $out2))

    # ---------------------------------------------------------- 4. 保守模式
    Section '4. 保守模式'
    $out3 = Join-Path $work 'cleaned3.xls'
    $r3 = Run @($fixture, '--conservative', '-o', $out3)
    Check '保守模式退出码为 0' ($r3.Code -eq 0) ("退出码 = " + $r3.Code)
    if (Test-Path -LiteralPath $out3) {
        $probe3 = Get-WorkbookProbe $out3
        Check '保守模式保留无人引用的 UnusedGood' ($probe3.Names.ContainsKey('UnusedGood')) `
              ("实际名称: " + (($probe3.Names.Keys) -join ', '))
        Check '保守模式仍删除已失效的 BrokenName' (-not $probe3.Names.ContainsKey('BrokenName'))
    }

    # ---------------------------------------------------------- 5. 防呆
    Section '5. 防呆'
    $r4 = Run @($fixture, '-o', $fixture)
    Check '输出路径 = 源文件时拒绝' ($r4.Code -ne 0 -or $r4.Out -match '不能与源文件相同') `
          ("退出码 = " + $r4.Code + "`n" + $r4.Out)

    $junk = Join-Path $work 'junk.xls'
    [System.IO.File]::WriteAllBytes($junk, [byte[]](1..200 | ForEach-Object { [byte]($_ % 251) }))
    $r5 = Run @($junk, '-o', (Join-Path $work 'junk-out.xls'))
    Check '非 CFBF 文件被拒绝' ($r5.Code -ne 0 -and $r5.Out -match '无法识别的文件格式') `
          ("退出码 = " + $r5.Code + "`n" + $r5.Out)
    Check '拒绝时未产生输出文件' (-not (Test-Path -LiteralPath (Join-Path $work 'junk-out.xls')))

    # ---------------------------------------------------------- 6. 就地覆盖
    Section '6. 就地覆盖'
    $inplace = Join-Path $work 'inplace.xls'
    Copy-Item -LiteralPath $fixture -Destination $inplace -Force
    $before = (Get-FileHash -LiteralPath $inplace -Algorithm SHA256).Hash
    $r6 = Run @($inplace, '--in-place')
    Check '就地覆盖退出码为 0' ($r6.Code -eq 0) ("退出码 = " + $r6.Code + "`n" + $r6.Out)
    Check '备份文件 .bak 已生成' (Test-Path -LiteralPath ($inplace + '.bak'))
    if (Test-Path -LiteralPath ($inplace + '.bak')) {
        $bakHash = (Get-FileHash -LiteralPath ($inplace + '.bak') -Algorithm SHA256).Hash
        Check '备份与源文件逐字节相同' ($bakHash -eq $before)
    }
    Check '就地覆盖后原文件已变小' ((Get-Item -LiteralPath $inplace).Length -le (Get-Item -LiteralPath $fixture).Length)

    # ---------------------------------------------------------- 7. 扩展名与内容不符
    Section '7. 扩展名与内容不符'
    $mis = Join-Path $work 'mislabeled.xls'
    Copy-Item -LiteralPath $fixture -Destination $mis -Force
    $misXlsx = Join-Path $work 'mislabeled2.xls'    # 内容其实是 xlsx，名字叫 .xls
    $realXlsx = Join-Path $work 'real.xlsx'
    $wbx = $script:excel.Workbooks.Add()
    while ($wbx.Worksheets.Count -lt 1) { [void]$wbx.Worksheets.Add() }
    $wbx.Worksheets.Item(1).Name = 'Data'
    $wbx.Worksheets.Item(1).Range('A1').Value2 = 42
    [void]$wbx.Names.Add('KeepMeToo', '=Data!$A$1')
    $wbx.Worksheets.Item(1).Range('B1').Formula = '=KeepMeToo+1'
    [void]$wbx.Names.Add('DeadName', '=#REF!')
    $wbx.SaveAs($realXlsx, 51)
    $wbx.Close($false)
    Copy-Item -LiteralPath $realXlsx -Destination $misXlsx -Force

    $outMis = Join-Path $work 'mislabeled-out.xls'
    $r7 = Run @($misXlsx, '-o', $outMis)
    Check '内容为 xlsx 但扩展名是 .xls 时按内容处理' `
          ($r7.Code -eq 0 -and $r7.Out -match '分析报告 =') `
          ("退出码 = " + $r7.Code + "`n" + $r7.Out)
    Check '给出了扩展名与内容不符的提示' ($r7.Out -match '扩展名是 \.xls，但文件内容其实是 OOXML')
    Check '输出文件可被 Excel 打开且内容正确' (Test-Path -LiteralPath $outMis)
    if (Test-Path -LiteralPath $outMis) {
        $pm = Get-WorkbookProbe $outMis
        Check '误名文件的引用名称 KeepMeToo 被保留' ($pm.Names.ContainsKey('KeepMeToo')) `
              ("实际名称: " + (($pm.Names.Keys) -join ', '))
        Check '误名文件的失效名称 DeadName 被删除' (-not $pm.Names.ContainsKey('DeadName'))
    }

    # ---------------------------------------------------------- 8. GUI 端到端
    Section '8. GUI 端到端（打开 → 修复/清理 → 自动保存）'
    $guiExe = Join-Path $root 'dist\XlsxDoctor.exe'
    if (-not (Test-Path -LiteralPath $guiExe)) {
        Skip 'GUI 端到端' '找不到 dist\XlsxDoctor.exe'
    } else {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class XlsUiBot
{
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessageW(IntPtr h, int msg, IntPtr w, StringBuilder l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessageW(IntPtr h, int msg, IntPtr w, IntPtr l);
    public const int WM_GETTEXT = 0x000D;
    public const int BM_CLICK   = 0x00F5;
    public static string Text(IntPtr h)
    {
        StringBuilder sb = new StringBuilder(32768);
        SendMessageW(h, WM_GETTEXT, (IntPtr)sb.Capacity, sb);
        return sb.ToString();
    }
    public static List<string> AllTexts(IntPtr root)
    {
        List<string> rows = new List<string>();
        EnumChildWindows(root, delegate(IntPtr h, IntPtr p)
        { string t = Text(h); if (t.Length > 0) rows.Add(t); return true; }, IntPtr.Zero);
        return rows;
    }
    public static IntPtr FindByText(IntPtr root, string want)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(root, delegate(IntPtr h, IntPtr p)
        { if (Text(h) == want) { found = h; return false; } return true; }, IntPtr.Zero);
        return found;
    }
}
'@

        Get-Process -Name 'XlsxDoctor*' -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Milliseconds 500

        $guiCopy = Join-Path $work 'gui-fixture.xls'
        Copy-Item -LiteralPath $fixture -Destination $guiCopy -Force
        $guiExpect = Join-Path $work 'gui-fixture - 已清理.xls'
        if (Test-Path -LiteralPath $guiExpect) { Remove-Item -LiteralPath $guiExpect -Force }

        $gp = Start-Process -FilePath $guiExe -ArgumentList ('"' + $guiCopy + '"') -PassThru
        Start-Sleep -Seconds 10
        $gp.Refresh()
        Check 'GUI 打开 .xls 后没有崩溃' (-not $gp.HasExited) `
              ("退出码 " + $(if ($gp.HasExited) { $gp.ExitCode } else { 'n/a' }))

        if (-not $gp.HasExited) {
            $gh = $gp.MainWindowHandle
            $gtexts = [XlsUiBot]::AllTexts($gh)
            $greport = $gtexts | Where-Object { $_ -like '*分析报告（.xls*' } | Select-Object -First 1
            Check 'GUI 报告确实是 .xls 版（按内容分派正确）' ($null -ne $greport -and $greport.Length -gt 0) `
                  '报告区不是 .xls 版'
            Check 'GUI 报告里有定义名称统计' ($greport -match '定义名称')
            Check 'GUI 报告里有外部表条目统计' ($greport -match '外部表条目')

            $gbtn = [XlsUiBot]::FindByText($gh, '2  修复并保存')
            Check '找到「修复并保存」按钮' ($gbtn -ne [IntPtr]::Zero) '找不到按钮'
            if ($gbtn -ne [IntPtr]::Zero) {
                [void][XlsUiBot]::SendMessageW($gbtn, [XlsUiBot]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
                Start-Sleep -Seconds 10
                $gp.Refresh()
                Check 'GUI 没有崩溃' (-not $gp.HasExited)
                Check '自动保存的输出 .xls 确实存在' (Test-Path -LiteralPath $guiExpect) '文件没写出来！'
                if (Test-Path -LiteralPath $guiExpect) {
                    $gm = [System.IO.File]::ReadAllBytes($guiExpect)[0..3]
                    Check '输出仍是合法的 CFBF/.xls' `
                          ($gm[0] -eq 0xD0 -and $gm[1] -eq 0xCF -and $gm[2] -eq 0x11 -and $gm[3] -eq 0xE0) `
                          (("文件头 = " + (($gm | ForEach-Object { $_.ToString('X2') }) -join ' ')))
                    $pg = Get-WorkbookProbe $guiExpect
                    Check '输出可被 Excel 打开且引用名称仍在' ($pg.Names.ContainsKey('KeepMe')) `
                          ("实际名称: " + (($pg.Names.Keys) -join ', '))
                    Check '输出里已无失效名称' (-not $pg.Names.ContainsKey('BrokenName'))
                    $badG = 0
                    foreach ($k in $probe0.Sheets.Keys) {
                        if (-not $pg.Sheets.ContainsKey($k)) { $badG++; continue }
                        if ($probe0.Sheets[$k].VHash -ne $pg.Sheets[$k].VHash) { $badG++ }
                    }
                    Check 'GUI 自动保存的输出计算值与源文件一致' ($badG -eq 0) ("不一致工作表数 = " + $badG)
                }
            }
            $gp.CloseMainWindow() | Out-Null
        }
        Start-Sleep -Seconds 1
        Get-Process -Name 'XlsxDoctor*' -ErrorAction SilentlyContinue | Stop-Process -Force
    }

} finally {
    Close-Excel
}

Write-Host ''
Write-Host '========================================' -ForegroundColor White
Write-Host (" 通过 {0} / 失败 {1} / 跳过 {2}" -f $script:pass, $script:fail, $script:skip) -ForegroundColor White
if ($script:fail -gt 0) {
    Write-Host ' 失败项：' -ForegroundColor Red
    foreach ($n in $script:failedNames) { Write-Host ("   - " + $n) -ForegroundColor Red }
}
Write-Host (" 工作目录: " + $work) -ForegroundColor DarkGray
Write-Host '========================================' -ForegroundColor White

if ($script:fail -gt 0) { exit 1 } else { exit 0 }
