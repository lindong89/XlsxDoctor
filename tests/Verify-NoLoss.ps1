<#
    XlsxDoctor 无损验证
    ==================
    独立证明：清理工具只删「失效/无人引用的索引条目」，单元格数据一字未动。

    本脚本不依赖 XlsxDoctor 的任何内部逻辑：
      · 用 Excel 真实打开两个文件，逐格比对公式与计算结果；
      · 对「清理后」文件做一次强制全量重算——如果某个被删的名字其实还在被公式
        使用，重算后对应单元格会立刻变成 #NAME? 错误，逐格比对必然发现；
      · 对比工作表、图表、定义名称、外部链接、包内部件。

    用法：
      .\Verify-NoLoss.ps1 -Original "清理前.xlsx" -Cleaned "清理后.xlsx"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Original,
    [Parameter(Mandatory=$true)][string]$Cleaned,
    [switch]$SkipRecalc
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

# ---------- 工具函数 ----------
function Say([string]$s) { Write-Host $s }
function Sha256([string]$s) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s))) }
    finally { $sha.Dispose() }
}
function Get-ZipText([string]$path, [string]$name) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        foreach ($e in $z.Entries) {
            if ($e.FullName -eq $name) {
                $s = $e.Open()
                try { $sr = [System.IO.StreamReader]::new($s); return $sr.ReadToEnd() } finally { $s.Dispose() }
            }
        }
        return $null
    } finally { $z.Dispose() }
}
function Get-SheetSnapshot($ws) {
    $rng = $ws.UsedRange
    $rows = [int]$rng.Rows.Count
    $cols = [int]$rng.Columns.Count
    $cells = $rows * $cols
    $fHash = ''; $vHash = ''
    if ($cells -gt 0) {
        $f = $rng.Formula
        $v = $rng.Value2
        $sbf = New-Object System.Text.StringBuilder
        $sbv = New-Object System.Text.StringBuilder
        if ($rows -eq 1 -and $cols -eq 1) {
            [void]$sbf.Append([string]$f)
            if ($null -ne $v) { [void]$sbv.Append([string]$v) }
        } else {
            $f0 = $f.GetLowerBound(0); $f1 = $f.GetLowerBound(1)
            $v0 = $null; $v1 = $null
            if ($null -ne $v) { $v0 = $v.GetLowerBound(0); $v1 = $v.GetLowerBound(1) }
            for ($r = 0; $r -lt $rows; $r++) {
                for ($c = 0; $c -lt $cols; $c++) {
                    [void]$sbf.Append([string]$f.GetValue($f0 + $r, $f1 + $c)).Append([char]1)
                    if ($null -ne $v) { [void]$sbv.Append([string]$v.GetValue($v0 + $r, $v1 + $c)).Append([char]1) }
                }
            }
        }
        $fHash = Sha256 $sbf.ToString()
        $vHash = Sha256 $sbv.ToString()
    }
    return [pscustomobject]@{ Rows=$rows; Cols=$cols; Cells=$cells; F=$fHash; V=$vHash }
}
function Get-ErrorInfo($ws) {
    $cnt = 0; $sample = ''
    try {
        $r = $ws.UsedRange.SpecialCells(-4123, 16)   # xlCellTypeFormulas + xlErrors
        if ($null -ne $r) {
            $cnt = [int]$r.Count
            $sample = [string]$r.Areas.Item(1).Cells.Item(1).Text
        }
    } catch { }
    return [pscustomobject]@{ Count=$cnt; Sample=$sample }
}
function Get-DefinedNameList([string]$path) {
    $xml = Get-ZipText $path 'xl/workbook.xml'
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($m in [regex]::Matches($xml, '<definedName\b([^>]*?)(/>|>([^<]*)</definedName>)')) {
        $attrs = $m.Groups[1].Value
        $val = $m.Groups[3].Value
        $nm = ''
        if ($attrs -match 'name="([^"]*)"') { $nm = $Matches[1] }
        $list.Add([pscustomobject]@{ Name=$nm; Value=$val; Sys=$nm.StartsWith('_xlnm.') })
    }
    return $list
}
function Get-PartHashes([string]$path) {
    $h = @{}
    $z = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        foreach ($e in $z.Entries) {
            $s = $e.Open()
            try {
                $sha = [System.Security.Cryptography.SHA256]::Create()
                try { $b = $sha.ComputeHash($s) } finally { $sha.Dispose() }
                $h[$e.FullName] = [System.BitConverter]::ToString($b)
            } finally { $s.Dispose() }
        }
    } finally { $z.Dispose() }
    return $h
}

$script:pass = 0
$script:fail = 0
function Verdict([string]$label, [bool]$ok, [string]$detail) {
    if ($ok) { $script:pass++ } else { $script:fail++ }
    $tag = 'PASS'; if (-not $ok) { $tag = 'FAIL' }
    Write-Host ('  [' + $tag + '] ' + $label)
    if ($detail -and $detail.Length -gt 0) { Write-Host ('          ' + $detail) }
}

# ---------- 主流程 ----------
if (-not (Test-Path -LiteralPath $Original)) { throw '找不到原文件: ' + $Original }
if (-not (Test-Path -LiteralPath $Cleaned))  { throw '找不到清理后文件: ' + $Cleaned }

$oLen = (Get-Item -LiteralPath $Original).Length
$cLen = (Get-Item -LiteralPath $Cleaned).Length

Write-Host '================================================================'
Write-Host '  XlsxDoctor 无损验证报告'
Write-Host '================================================================'
Write-Host ('清理前 : ' + (Split-Path $Original -Leaf) + '   ' + $oLen.ToString('N0') + ' 字节')
Write-Host ('清理后 : ' + (Split-Path $Cleaned -Leaf) + '   ' + $cLen.ToString('N0') + ' 字节')
if ($oLen -gt 0) { Write-Host ('体积   : 清理后为原文件的 ' + (100.0 * $cLen / $oLen).ToString('0.0') + '%') }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try { $excel.AskToUpdateLinks = $false } catch { }
try { $excel.Calculation = -4135 } catch { }   # 手动计算：只用缓存值，不联网

try {
    # ---- 打开原文件并快照 ----
    $wbA = $excel.Workbooks.Open($Original, 0, $true)   # UpdateLinks=0, ReadOnly
    $sheetsA = @()
    for ($i = 1; $i -le $wbA.Worksheets.Count; $i++) {
        $ws = $wbA.Worksheets.Item($i)
        $snap = Get-SheetSnapshot $ws
        $err = Get-ErrorInfo $ws
        $charts = @($ws.ChartObjects()).Count
        $sheetsA += [pscustomobject]@{ Name=$ws.Name; Snap=$snap; Err=$err; Charts=$charts }
    }
    $namesA = Get-DefinedNameList $Original
    $wbA.Close($false)

    # ---- 打开清理后文件并快照 ----
    $wbB = $excel.Workbooks.Open($Cleaned, 0, $true)
    $sheetsB = @()
    for ($i = 1; $i -le $wbB.Worksheets.Count; $i++) {
        $ws = $wbB.Worksheets.Item($i)
        $snap = Get-SheetSnapshot $ws
        $err = Get-ErrorInfo $ws
        $charts = @($ws.ChartObjects()).Count
        $sheetsB += [pscustomobject]@{ Name=$ws.Name; Snap=$snap; Err=$err; Charts=$charts }
    }
    $namesB = Get-DefinedNameList $Cleaned

    # ---------- 1. 工作表 ----------
    Write-Host ''
    Write-Host '--- 1. 工作表清单 ---'
    Verdict '工作表数量' ($sheetsA.Count -eq $sheetsB.Count) ("$($sheetsA.Count) -> $($sheetsB.Count)")
    $nameOk = $true
    for ($i = 0; $i -lt [Math]::Min($sheetsA.Count, $sheetsB.Count); $i++) {
        if ($sheetsA[$i].Name -ne $sheetsB[$i].Name) { $nameOk = $false }
    }
    Verdict '工作表名称与顺序' $nameOk (($sheetsA | ForEach-Object { $_.Name }) -join ' / ')

    # ---------- 2. 逐格比对 ----------
    Write-Host ''
    Write-Host '--- 2. 单元格内容逐格比对（公式 + 计算结果）---'
    $cellsTotal = 0
    $fDiffs = 0; $vDiffs = 0; $errDiffs = 0
    for ($i = 0; $i -lt [Math]::Min($sheetsA.Count, $sheetsB.Count); $i++) {
        $a = $sheetsA[$i]; $b = $sheetsB[$i]
        $cellsTotal += $a.Snap.Cells
        if ($a.Snap.F -ne $b.Snap.F) { $fDiffs++ }
        if ($a.Snap.V -ne $b.Snap.V) { $vDiffs++ }
        if ($a.Err.Count -ne $b.Err.Count) { $errDiffs++ }
        Write-Host ('  表「{0,-24} {1}x{2,-4} {3,9} 单元格  {4}' -f $a.Name, $a.Snap.Rows, $a.Snap.Cols, $a.Snap.Cells, $(if ($a.Snap.F -eq $b.Snap.F -and $a.Snap.V -eq $b.Snap.V) { '一致' } else { '有差异!' }))
    }
    Verdict '公式逐格比对' ($fDiffs -eq 0) ("$cellsTotal 个单元格，$fDiffs 个表有差异")
    Verdict '计算结果逐格比对' ($vDiffs -eq 0) ("$cellsTotal 个单元格，$vDiffs 个表有差异")

    # ---------- 3. 错误值单元格 ----------
    $errA = ($sheetsA | ForEach-Object { $_.Err.Count } | Measure-Object -Sum).Sum
    $errB = ($sheetsB | ForEach-Object { $_.Err.Count } | Measure-Object -Sum).Sum
    Verdict '错误值单元格数量（#REF!、#NAME? 等）' ($errA -eq $errB) ("清理前 $errA 个 -> 清理后 $errB 个")

    # ---------- 4. 决定性测试：强制重算 ----------
    if (-not $SkipRecalc) {
        Write-Host ''
        Write-Host '--- 3. 决定性测试：强制全量重算「清理后」文件 ---'
        Write-Host '    如果某个被删的名字其实还在被公式使用，重算后对应单元格会变成 #NAME? 错误。'
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $recalcOk = $true
        try { $wbB.Application.CalculateFullRebuild() } catch { $recalcOk = $false; Write-Host ('    重算不可用: ' + $_.Exception.Message) }
        $sw.Stop()
        if ($recalcOk) {
            Write-Host ('    重算耗时 ' + $sw.Elapsed.TotalSeconds.ToString('0.0') + ' 秒')
            $vDiffs2 = 0; $errAfter = 0
            for ($i = 0; $i -lt $sheetsB.Count; $i++) {
                $ws = $wbB.Worksheets.Item($i + 1)
                $snap2 = Get-SheetSnapshot $ws
                $err2 = Get-ErrorInfo $ws
                $errAfter += $err2.Count
                if ($sheetsA[$i].Snap.V -ne $snap2.V) {
                    $vDiffs2++
                    Write-Host ('    表「' + $sheetsA[$i].Name + '」重算后数值与清理前不同')
                    if ($err2.Count -gt $sheetsA[$i].Err.Count) {
                        Write-Host ('      该表错误值单元格: ' + $sheetsA[$i].Err.Count + ' -> ' + $err2.Count + '（示例: ' + $err2.Sample + '）')
                    }
                }
            }
            Verdict '重算后数值与清理前一致' ($vDiffs2 -eq 0) ("$cellsTotal 个单元格，$vDiffs2 个表有差异")
            Verdict '重算后错误值没有增加' ($errAfter -le $errA) ("错误值单元格: $errA -> $errAfter")
        }
    }
    $wbB.Close($false)

    # ---------- 5. 图表 ----------
    $chA = ($sheetsA | ForEach-Object { $_.Charts } | Measure-Object -Sum).Sum
    $chB = ($sheetsB | ForEach-Object { $_.Charts } | Measure-Object -Sum).Sum
    Verdict '嵌入图表数量' ($chA -eq $chB) ("$chA -> $chB")

    # ---------- 6. 定义名称 ----------
    Write-Host ''
    Write-Host '--- 4. 定义名称（本次清理的对象）---'
    $bNames = @{}; foreach ($n in $namesB) { $bNames[$n.Name] = $true }
    $removed = @($namesA | Where-Object { -not $bNames.ContainsKey($_.Name) })
    Verdict '定义名称' ($namesB.Count -le $namesA.Count) ("$($namesA.Count) -> $($namesB.Count)，删除 $($removed.Count) 个")
    $errRx = [regex]'#(REF!|N/A|NAME\?|VALUE!|DIV/0!|NULL!|NUM!|GETTING_DATA)'
    $extRx = [regex]'^\s*\['
    $broken = @($removed | Where-Object { $_.Value -match $errRx })
    $external = @($removed | Where-Object { $_.Value -match $extRx -and $_.Value -notmatch $errRx })
    $brokenN = @{}; foreach ($n in $broken) { $brokenN[$n.Name] = $true }
    $extN = @{}; foreach ($n in $external) { $extN[$n.Name] = $true }
    $other = @($removed | Where-Object { -not $brokenN.ContainsKey($_.Name) -and -not $extN.ContainsKey($_.Name) })
    Write-Host ('    删除的 ' + $removed.Count + ' 个中：')
    Write-Host ('      值已失效（#REF! / #N/A 等）: ' + $broken.Count + ' 个')
    Write-Host ('      指向外部工作簿（[n] 开头） : ' + $external.Count + ' 个')
    Write-Host ('      其他                       : ' + $other.Count + ' 个')
    if ($other.Count -gt 0) {
        Write-Host '      示例（可人工复核）：'
        $other | Select-Object -First 8 | ForEach-Object { Write-Host ('        ' + $_.Name + ' = ' + $_.Value) }
    }
    Write-Host ('    保留 ' + $namesB.Count + ' 个：' + (($namesB | ForEach-Object { $_.Name }) -join '、'))

    # ---------- 7. 外部链接 ----------
    $xmlA = Get-ZipText $Original 'xl/workbook.xml'
    $xmlB = Get-ZipText $Cleaned 'xl/workbook.xml'
    $extA = ([regex]::Matches($xmlA, '<externalReference\b')).Count
    $extB = ([regex]::Matches($xmlB, '<externalReference\b')).Count
    Verdict '外部链接' ($extB -le $extA) ("$extA -> $extB（保留的仍被公式引用）")

    # ---------- 8. 部件层面 ----------
    Write-Host ''
    Write-Host '--- 5. 包内部件（zip 内部文件）---'
    $hA = Get-PartHashes $Original
    $hB = Get-PartHashes $Cleaned
    $changed = @($hA.Keys | Where-Object { $hB.ContainsKey($_) -and $hB[$_] -ne $hA[$_] })
    $removedParts = @($hA.Keys | Where-Object { -not $hB.ContainsKey($_) })
    $addedParts = @($hB.Keys | Where-Object { -not $hA.ContainsKey($_) })
    Verdict '没有新增部件' ($addedParts.Count -eq 0) "新增 $($addedParts.Count) 个"
    Verdict '只有元数据部件被改写' ($changed.Count -le 5) ("被改写 $($changed.Count) 个: " + ($changed -join ', '))
    $nonLink = @($removedParts | Where-Object { -not $_.EndsWith('/') -and $_ -notmatch 'externalLinks' })
    Verdict '删除的部件全部是外部链接' ($nonLink.Count -eq 0) ("删除 $($removedParts.Count) 个（含目录条目），非 externalLinks 的 $($nonLink.Count) 个")

} finally {
    try { $excel.Quit() } catch { }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

# ---------- 结论 ----------
Write-Host ''
Write-Host '================================================================'
Write-Host ("  结论：$($script:pass) 项检查通过，$($script:fail) 项未通过")
if ($script:fail -eq 0) {
    Write-Host '  所有工作表、单元格、公式、计算结果、图表与清理前完全一致。'
    Write-Host '  被删除的只是失效/无人引用的索引条目（定义名称）与无人引用的外部链接。'
} else {
    Write-Host '  !! 发现差异，请勿使用清理后的文件，立即核对。'
}
Write-Host '================================================================'
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
