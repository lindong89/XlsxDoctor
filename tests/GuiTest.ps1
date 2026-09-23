<#
    XlsxDoctor 图形界面端到端测试
    真的启动 GUI、真的点「修复并保存」按钮，验证自动保存确实落盘。

    用法：
        .\tests\GuiTest.ps1 -Src <一个含垃圾的 xlsx>
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Src,
    [string]$Exe
)

$ErrorActionPreference = 'Stop'

# 任何异常都必须顺手清掉 GUI 进程：残留实例会锁住 exe，导致下次 build.ps1 编译失败
trap {
    Write-Host ("测试异常中断: " + $_.Exception.Message) -ForegroundColor Red
    Get-Process -Name 'XlsxDoctor*' -ErrorAction SilentlyContinue | Stop-Process -Force
    exit 1
}

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Definition)
if (-not $Exe) { $Exe = Join-Path $root 'dist\XlsxDoctor.exe' }
if (-not (Test-Path -LiteralPath $Exe)) { throw "找不到 $Exe，请先运行 .\build.ps1" }
$Exe = (Resolve-Path -LiteralPath $Exe).Path
if (-not (Test-Path -LiteralPath $Src)) { throw "找不到源文件 $Src" }
if (Test-Path -LiteralPath $Src -PathType Container) {
    # 传目录进来是很容易犯的错：目录会被复制成一个叫 src 的目录，
    # GUI 发现参数不是文件就退回命令行模式，然后以退出码 1 结束，现象很迷惑。
    throw "-Src 需要一个**夹具文件**，不是目录。你传的是目录: $Src`n例如: -Src .\tests\fixtures\junk.xlsx"
}

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public class UiBot
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessageW(IntPtr h, int msg, IntPtr w, StringBuilder l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessageW(IntPtr h, int msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);

    public const int WM_GETTEXT = 0x000D;
    public const int BM_CLICK   = 0x00F5;

    public static string Text(IntPtr h)
    {
        StringBuilder sb = new StringBuilder(8192);
        SendMessageW(h, WM_GETTEXT, (IntPtr)sb.Capacity, sb);
        return sb.ToString();
    }

    public static IntPtr FindByText(IntPtr root, string want)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(root, delegate(IntPtr h, IntPtr p)
        {
            if (Text(h) == want) { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static List<string> AllTexts(IntPtr root)
    {
        List<string> rows = new List<string>();
        EnumChildWindows(root, delegate(IntPtr h, IntPtr p)
        {
            string t = Text(h);
            if (t.Length > 0) rows.Add(t);
            return true;
        }, IntPtr.Zero);
        return rows;
    }
}
'@

# 在临时目录里放一份副本，别动原文件
$work = Join-Path $env:TEMP ('xlsxdoctor-gui-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work | Out-Null
$copy = Join-Path $work (Split-Path -Leaf $Src)
Copy-Item -LiteralPath $Src -Destination $copy -Force

$base = [System.IO.Path]::GetFileNameWithoutExtension($copy)
$expect = Join-Path $work ($base + ' - 已清理.xlsx')

$pass = 0; $fail = 0

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntry([string]$zipPath, [string]$entryName) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $e = $z.Entries | Where-Object { $_.FullName -eq $entryName }
        if (-not $e) { return [string]::Empty }
        $s = $e.Open()
        try {
            $sr = [System.IO.StreamReader]::new($s)
            try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
        } finally { $s.Dispose() }
    } finally { $z.Dispose() }
}
function Check([string]$n, [bool]$ok, [string]$d) {
    if ($ok) { $script:pass++; Write-Host ("  [PASS] {0}" -f $n) -ForegroundColor Green }
    else     { $script:fail++; Write-Host ("  [FAIL] {0}" -f $n) -ForegroundColor Red; if ($d) { Write-Host ("         {0}" -f $d) -ForegroundColor DarkRed } }
}

Write-Host ''
Write-Host 'XlsxDoctor 图形界面端到端测试' -ForegroundColor White
Write-Host ("exe     : {0}" -f $Exe)
Write-Host ("副本    : {0}" -f $copy)
Write-Host ("期望输出: {0}" -f $expect)
Write-Host ''

# 先清掉可能残留的旧实例
Get-Process -Name 'XlsxDoctor*' -ErrorAction SilentlyContinue | Stop-Process -Force

$proc = Start-Process -FilePath $Exe -ArgumentList ('"' + $copy + '"') -PassThru
Write-Host ("已启动 GUI，PID {0}，等待分析…" -f $proc.Id)
Start-Sleep -Seconds 12

$proc.Refresh()
if ($proc.HasExited) { Write-Host ("!! GUI 提前退出，退出码 {0}" -f $proc.ExitCode) -ForegroundColor Red; exit 1 }

$h = $proc.MainWindowHandle
Check 'GUI 窗口已创建' ($h -ne [IntPtr]::Zero) '没有主窗口'

$texts = [UiBot]::AllTexts($h)
$report = $texts | Where-Object { $_ -like '*分析报告*' } | Select-Object -First 1
Check '打开文件后自动完成分析' ($null -ne $report -and $report.Length -gt 0) '报告区是空的'
Check '报告里有定义名称统计' ($report -match '定义名称') '报告内容不对'
Check '报告里有外部链接统计' ($report -match '外部链接') '报告内容不对'
Check '状态栏提示可修复' (($texts | Where-Object { $_ -like '*点「修复并保存」执行*' }).Count -gt 0) '状态栏没提示'

# 点「修复并保存」
$btn = [UiBot]::FindByText($h, '2  修复并保存')
Check '找到「修复并保存」按钮' ($btn -ne [IntPtr]::Zero) '找不到按钮'
if ($btn -eq [IntPtr]::Zero) { $proc.Kill(); exit 1 }

Write-Host '  点击「修复并保存」…' -ForegroundColor DarkGray
[void][UiBot]::SendMessageW($btn, [UiBot]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
Start-Sleep -Seconds 12

$proc.Refresh()
Check 'GUI 没有崩溃' (-not $proc.HasExited) '进程退出了'

$texts = [UiBot]::AllTexts($h)
$report = $texts | Where-Object { $_ -like '*分析报告*' } | Select-Object -First 1
Check '报告显示包校验通过' ($report -match '包校验\s*:\s*通过') '校验没通过'
Check '报告显示输出文件路径' ($report -match [regex]::Escape($expect)) '报告里的输出路径不对'
Check '状态栏显示完成' (($texts | Where-Object { $_ -like '*完成：*' }).Count -gt 0) '状态栏没显示完成'

Check '自动保存的输出文件确实存在' (Test-Path -LiteralPath $expect) '文件没写出来！'

if (Test-Path -LiteralPath $expect) {
    $inSize  = (Get-Item -LiteralPath $copy).Length
    $outSize = (Get-Item -LiteralPath $expect).Length
    Check '输出文件比源文件小' ($outSize -lt $inSize) ("源 $inSize -> 出 $outSize")

    $srcWb = Read-ZipEntry $copy 'xl/workbook.xml'
    $wb    = Read-ZipEntry $expect 'xl/workbook.xml'

    $srcNames = ([regex]::Matches($srcWb, '<definedName\b')).Count
    $outNames = ([regex]::Matches($wb, '<definedName\b')).Count
    Check '定义名称数量大幅下降' ($outNames -lt $srcNames) ("源 $srcNames -> 出 $outNames")
    Check '输出里已无失效的名称值' (-not ($wb -match '#REF!|#N/A')) '还有 #REF!/#N/A'
    # 注意：被公式引用的外部链接是必须保留的，所以这里只能比"下降"，不能要求归零
    $srcLinks = ([regex]::Matches($srcWb, '<externalReference\b')).Count
    $outLinks = ([regex]::Matches($wb, '<externalReference\b')).Count
    Check '外部链接引用数量下降' ($outLinks -lt $srcLinks) ("源 $srcLinks -> 出 $outLinks")
}

$proc.CloseMainWindow() | Out-Null
Start-Sleep -Seconds 1
# 兜底：按名字扫一遍强制结束。残留的 GUI 会锁住 exe，导致下次 build.ps1 报 CS0016
Get-Process -Name 'XlsxDoctor*' -ErrorAction SilentlyContinue | Stop-Process -Force

Write-Host ''
Write-Host ("通过 {0}   失败 {1}" -f $pass, $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
Write-Host ("工作目录: {0}" -f $work)
Write-Host ''
if ($fail -eq 0) { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue; exit 0 }
exit 1
